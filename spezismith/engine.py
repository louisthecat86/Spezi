"""
SpeziSmith - Kern-Engine

Liest Spezifikations-Vorlagen (xlsx), stellt jedes Feld als waehlbare
Ersetzungsregel bereit und schreibt geaenderte Werte per ZIP/XML-Surgery
zurueck. Dadurch bleiben eingebettete Bilder (Firmenlogo), VML-Grafiken,
ActiveX-Steuerelemente, Formatierungen und Druckeinstellungen erhalten; bei
Zeilenaenderungen werden VML-Steuerelemente passend mitverschoben.
"""

from __future__ import annotations

import os
import copy
import difflib
import posixpath
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Callable, Iterable, Optional

import openpyxl

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

# Markerzellen zur Spracherkennung (A3 = Dokumenttitel)
LANG_MARKERS = {
    "de": ("produktspezifizierung", "artikelnummer", "artikelbezeichnung"),
    "en": ("product-specification", "product specification", "article number", "article"),
}


# ---------------------------------------------------------------------------
# Zell-Koordinaten
# ---------------------------------------------------------------------------

def col_to_idx(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def idx_to_col(idx: int) -> str:
    col = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        col = chr(65 + rem) + col
    return col


def split_coord(coord: str) -> tuple[str, int]:
    m = re.match(r"^\$?([A-Za-z]+)\$?(\d+)$", coord)
    if not m:
        raise ValueError(f"Ungueltiger Zellbezug: {coord}")
    return m.group(1).upper(), int(m.group(2))


def coord_key(coord: str) -> tuple[int, int]:
    col, row = split_coord(coord)
    return row, col_to_idx(col)


def norm_text(value) -> str:
    """Normalisiert Zellwerte fuer Vergleiche."""
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"[ \t]+", " ", text).strip()


# ---------------------------------------------------------------------------
# Datenmodell
# ---------------------------------------------------------------------------

@dataclass
class Field:
    """Ein waehlbares Feld der Spezifikation."""
    coord: str            # Zellbezug (Ankerzelle bei verbundenen Zellen)
    row: int
    col: int
    value: str            # aktueller Wert ("" wenn leer)
    label: str            # Kontext-Beschriftung (Nachbarzelle links/oberhalb)
    is_input: bool = False  # True = leeres Eingabefeld (Kandidat)

    @property
    def display_value(self) -> str:
        return self.value if self.value else "(leer)"


@dataclass
class Rule:
    """Eine Ersetzungsregel fuer genau eine Zelle."""
    coord: str
    new: str
    mode: str = "set"          # "set" = ganze Zelle, "replace" = Teiltext
    find: str = ""             # Suchtext fuer mode="replace"
    expected: str = ""         # Vorlagenwert fuer Verifikation/Inhaltssuche
    label: str = ""
    anchor_text: str = ""      # Beschriftung neben/ueber dem Feld (Verschiebungs-Anker)
    anchor_coord: str = ""     # Koordinate der Beschriftung in der Vorlage
    context_above: list = field(default_factory=list)  # Texte darueber (nah -> fern)
    line_start: int = -1       # mehrzeilig: erste geaenderte Zeile (-1 = unbekannt)
    line_count: int = 0        # mehrzeilig: Anzahl der geaenderten Zeilen
    allow_insert: bool = False  # fehlendes Feld als neue Zeile einfuegen


# Kopf-/Fußzeilen werden als Pseudo-Koordinaten adressiert: "HF:oddFooter.R"
HF_PREFIX = "HF:"
HF_PARTS = ("oddHeader", "oddFooter", "evenHeader", "evenFooter",
            "firstHeader", "firstFooter")


def split_hf_coord(coord: str) -> tuple[str, str]:
    """'HF:oddFooter.R' -> ('oddFooter', 'R')."""
    body = coord[len(HF_PREFIX):]
    part, _, section = body.rpartition(".")
    if part not in HF_PARTS or section.upper() not in ("L", "C", "R"):
        raise ValueError(f"Ungueltige Kopf-/Fußzeilen-Adresse: {coord}")
    return part, section.upper()


@dataclass
class CellResult:
    coord: str
    label: str
    old: str
    new: str
    status: str                # "geaendert" | "uebersprungen" | "fehler"
    message: str = ""


@dataclass
class FileResult:
    source: str
    target: str = ""
    ok: bool = False
    language: Optional[str] = None
    cells: list[CellResult] = field(default_factory=list)
    structure_notes: list[str] = field(default_factory=list)
    error: str = ""

    @property
    def changed_count(self) -> int:
        return sum(1 for c in self.cells if c.status == "geaendert")

    @property
    def skipped_count(self) -> int:
        return sum(1 for c in self.cells if c.status == "uebersprungen")


# ---------------------------------------------------------------------------
# Vorlagen-Analyse
# ---------------------------------------------------------------------------

def detect_language(path: str) -> Optional[str]:
    """Erkennt die Sprache einer Spezifikation anhand der Markerzellen."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        ws = wb.worksheets[0]
        probe = []
        for row in ws.iter_rows(min_row=1, max_row=8, max_col=6):
            for cell in row:
                if cell.value is not None:
                    probe.append(norm_text(cell.value).lower())
        text = " | ".join(probe)
    finally:
        wb.close()
    scores = {
        lang: sum(1 for marker in markers if marker in text)
        for lang, markers in LANG_MARKERS.items()
    }
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else None


def template_row_texts(path: str) -> dict[int, str]:
    """Erster Text je Zeile der Vorlage (Zeile -> Text); Grundlage der Einfuegeposition."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows: dict[int, str] = {}
        for row in wb.worksheets[0].iter_rows():
            for cell in row:
                if cell.value is not None and norm_text(cell.value):
                    rows[cell.row] = norm_text(cell.value)
                    break
        return rows
    finally:
        wb.close()


def context_above(rows: dict[int, str], row: int, limit: int = 6) -> list[str]:
    """Texte der Zeilen oberhalb von `row` (nah -> fern, ohne Doppelte)."""
    out: list[str] = []
    for r in sorted((r for r in rows if r < row), reverse=True):
        if rows[r] not in out:
            out.append(rows[r])
        if len(out) >= limit:
            break
    return out


def analyze_template(path: str, include_empty_inputs: bool = True) -> list[Field]:
    """
    Analysiert eine Vorlage und liefert alle waehlbaren Felder:
    - jede befuellte Zelle (Ankerzelle bei Verbundenen)
    - zusaetzlich leere Eingabefelder neben Beschriftungen (optional)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        max_row, max_col = ws.max_row, ws.max_column

        # Verbundene Zellen: Anker <-> Bereich
        anchor_of: dict[str, str] = {}
        span_of: dict[str, list[str]] = {}
        for rng in ws.merged_cells.ranges:
            anchor = f"{idx_to_col(rng.min_col)}{rng.min_row}"
            members = []
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    coord = f"{idx_to_col(c)}{r}"
                    anchor_of[coord] = anchor
                    members.append(coord)
            span_of[anchor] = members

        # Werte aller Anker-/Einzelzellen einsammeln
        values: dict[str, object] = {}
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                coord = cell.coordinate
                anchor = anchor_of.get(coord, coord)
                if cell.value is not None and norm_text(cell.value) != "":
                    values.setdefault(anchor, cell.value)

        def cell_text(coord: str) -> str:
            v = values.get(anchor_of.get(coord, coord))
            return norm_text(v) if v is not None else ""

        def infer_label(row: int, col: int, own: str) -> str:
            # naechste Textzelle links in derselben Zeile
            for c in range(col - 1, 0, -1):
                t = cell_text(f"{idx_to_col(c)}{row}")
                if t:
                    return t[:60]
            # sonst naechste Textzelle oberhalb in derselben Spalte
            for r in range(row - 1, 0, -1):
                t = cell_text(f"{idx_to_col(col)}{r}")
                if t:
                    return t[:60]
            return ""

        fields: list[Field] = []
        seen: set[str] = set()

        # 1) befuellte Zellen
        for anchor, value in values.items():
            col_s, row = split_coord(anchor)
            col = col_to_idx(col_s)
            own = norm_text(value)
            fields.append(Field(
                coord=anchor, row=row, col=col, value=own,
                label=infer_label(row, col, own), is_input=False,
            ))
            seen.add(anchor)

        # 2) leere Eingabefelder: leere Anker-/Einzelzellen mit Text links daneben
        if include_empty_inputs:
            for r in range(1, max_row + 1):
                for c in range(2, max_col + 1):
                    coord = f"{idx_to_col(c)}{r}"
                    anchor = anchor_of.get(coord, coord)
                    if anchor in seen or anchor != coord:
                        continue
                    left = cell_text(f"{idx_to_col(c - 1)}{r}")
                    if left:  # Beschriftung links vorhanden -> Eingabefeld
                        fields.append(Field(
                            coord=coord, row=r, col=c, value="",
                            label=left[:60], is_input=True,
                        ))
                        seen.add(coord)

        fields.sort(key=lambda f: (f.row, f.col))
        return fields
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# ZIP/XML-Surgery (format- und bild-sicheres Ersetzen)
# ---------------------------------------------------------------------------

_ROOT_START_RE = re.compile(r"<((?:[A-Za-z0-9_]+:)?[A-Za-z0-9_]+)((?:\s+[^<>]*)?)>")
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def _root_start_tag(xml_text: str) -> str:
    """Liefert den kompletten Start-Tag des Wurzelelements (hinter <?xml ... ?>)."""
    body = xml_text.split("?>", 1)[1] if "?>" in xml_text else xml_text
    m = _ROOT_START_RE.search(body)
    return m.group(0) if m else ""


def _register_namespaces(xml_text: str) -> list[tuple[str, str]]:
    """Registriert alle im Root deklarierten Namespaces bei ElementTree."""
    decls = re.findall(r'xmlns(?::([A-Za-z0-9_.]+))?="([^"]+)"', _root_start_tag(xml_text))
    for prefix, uri in decls:
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass
    return decls


def _restore_namespaces(xml_text: str, decls: list[tuple[str, str]]) -> str:
    """
    Stellt sicher, dass alle urspruenglichen Namespace-Deklarationen im Root
    vorhanden sind - insbesondere alle Praefixe aus mc:Ignorable, denn Excel
    meldet sonst einen XML-Fehler und startet die 'Reparatur'.
    """
    root_tag = _root_start_tag(xml_text)
    if not root_tag:
        raise ValueError("Wurzelelement im serialisierten XML nicht gefunden.")

    known = dict(decls)
    missing: list[tuple[str, str]] = []
    for prefix, uri in decls:
        marker = f'xmlns:{prefix}=' if prefix else "xmlns="
        if marker not in root_tag:
            missing.append((prefix, uri))

    new_root = root_tag
    if missing:
        inject = "".join(
            f' xmlns:{p}="{u}"' if p else f' xmlns="{u}"' for p, u in missing
        )
        new_root = root_tag[:-1] + inject + ">"

    # Praefixe in mc:Ignorable ohne bekannte URI herausloesen
    ign = re.search(r'(?:\w+:)?Ignorable="([^"]+)"', new_root)
    if ign:
        keep = []
        for prefix in ign.group(1).split():
            if prefix in known or f'xmlns:{prefix}=' in new_root:
                keep.append(prefix)
        if keep != ign.group(1).split():
            new_root = new_root.replace(
                ign.group(0),
                ign.group(0).replace(f'"{ign.group(1)}"', f'"{" ".join(keep)}"'),
            )

    return xml_text.replace(root_tag, new_root, 1)


def _validate_sheet_xml(xml_text: str) -> None:
    """Strikte Validierung, damit Excel die Datei nicht 'reparieren' muss."""
    ET.fromstring(xml_text.encode("utf-8"))  # Wohlgeformtheit
    root_tag = _root_start_tag(xml_text)
    attrs = re.findall(r'(xmlns(?::[A-Za-z0-9_.]+)?)="', root_tag)
    if len(attrs) != len(set(attrs)):
        raise ValueError("Doppelte xmlns-Attribute im Wurzelelement.")
    ign = re.search(r'(?:\w+:)?Ignorable="([^"]+)"', root_tag)
    if ign:
        for prefix in ign.group(1).split():
            if f'xmlns:{prefix}=' not in root_tag:
                raise ValueError(f"mc:Ignorable-Praefix '{prefix}' nicht deklariert.")


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings = []
    for si in root.findall(f"{{{MAIN_NS}}}si"):
        parts = [t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t")]
        strings.append("".join(parts))
    return strings


def _first_sheet_path(zf: zipfile.ZipFile) -> str:
    """Ermittelt den ZIP-Pfad des ersten Arbeitsblatts."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    sheet = wb_root.find(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet")
    if sheet is None:
        raise ValueError("Keine Tabelle in der Datei gefunden.")
    rid = sheet.get(f"{{{REL_NS}}}id")
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if rel.get("Id") == rid:
            target = rel.get("Target")
            if not target.startswith("/"):
                target = "xl/" + target
            return target.lstrip("/")
    raise ValueError("Tabellen-Relationship nicht gefunden.")


def _sheet_related_paths(
    zf: zipfile.ZipFile, sheet_path: str, relationship_type: str
) -> set[str]:
    rels_path = posixpath.join(
        posixpath.dirname(sheet_path), "_rels",
        posixpath.basename(sheet_path) + ".rels")
    try:
        rels_root = ET.fromstring(zf.read(rels_path))
    except KeyError:
        return set()
    paths = set()
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if not rel.get("Type", "").endswith(f"/{relationship_type}"):
            continue
        target = rel.get("Target", "")
        if target.startswith("/"):
            paths.add(target.lstrip("/"))
        else:
            paths.add(posixpath.normpath(
                posixpath.join(posixpath.dirname(sheet_path), target)))
    return paths


def _merge_anchor_map(sheet_root: ET.Element) -> dict[str, str]:
    mapping: dict[str, str] = {}
    merge_cells = sheet_root.find(f"{{{MAIN_NS}}}mergeCells")
    if merge_cells is None:
        return mapping
    for mc in merge_cells.findall(f"{{{MAIN_NS}}}mergeCell"):
        ref = mc.get("ref", "")
        if ":" not in ref:
            continue
        start, end = ref.split(":")
        c1, r1 = split_coord(start)
        c2, r2 = split_coord(end)
        for r in range(r1, r2 + 1):
            for ci in range(col_to_idx(c1), col_to_idx(c2) + 1):
                mapping[f"{idx_to_col(ci)}{r}"] = start
    return mapping


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    t = cell.get("t")
    if t == "s":
        v = cell.find(f"{{{MAIN_NS}}}v")
        if v is None or v.text is None:
            return ""
        idx = int(v.text)
        return shared[idx] if 0 <= idx < len(shared) else ""
    if t == "inlineStr":
        is_el = cell.find(f"{{{MAIN_NS}}}is")
        if is_el is None:
            return ""
        return "".join(x.text or "" for x in is_el.iter(f"{{{MAIN_NS}}}t"))
    v = cell.find(f"{{{MAIN_NS}}}v")
    return v.text if (v is not None and v.text is not None) else ""


# ---------------------------------------------------------------------------
# Kopf-/Fußzeile (Seitenkopf/-fuß der Spezifikation)
# ---------------------------------------------------------------------------

_HF_FONT_RE = re.compile(r'&"[^"]*"')
_HF_SIZE_RE = re.compile(r"&\d+(?:\.\d+)?")
_HF_COLOR_RE = re.compile(r"&[Kk][0-9A-Fa-f]{6}")


def hf_display_text(raw: str) -> str:
    """Lesbare Anzeige: Font-/Groessen-/Farb-Codes entfernen,
    Feld-Codes wie &P (Seite) und &N (Seiten gesamt) behalten."""
    text = _HF_FONT_RE.sub("", raw)
    text = _HF_COLOR_RE.sub("", text)
    text = _HF_SIZE_RE.sub("", text)
    return text


def _hf_split_sections(raw: str) -> tuple[dict[str, str], list[str]]:
    """Zerteilt '&L...&C...&R...' in Sektionen; merkt sich die Reihenfolge."""
    parts = re.split(r"&([LCR])", raw)
    sections: dict[str, str] = {}
    order: list[str] = []
    if parts[0]:
        sections["L"] = parts[0]  # Text ohne Marker = linker Teil
        order.append("L")
    for i in range(1, len(parts) - 1, 2):
        marker, text = parts[i], parts[i + 1]
        sections[marker] = sections.get(marker, "") + text
        if marker not in order:
            order.append(marker)
    return sections, order


def _hf_join_sections(sections: dict[str, str], order: list[str]) -> str:
    return "".join(f"&{m}{sections[m]}" for m in order)


def _hf_smart_replace(raw_section: str, new_text: str) -> tuple[str, str]:
    """
    Ersetzt den sichtbaren Text einer Kopf-/Fußzeilen-Sektion.
    Zeilenweiser Ersatz, solange die Zeilenzahl passt - so bleiben
    Format- und Feld-Codes (&P, &G, Schriftgroessen) exakt erhalten.
    Liefert (neuer_rohtext, warnung).
    """
    old_lines = [ln for ln in hf_display_text(raw_section).split("\n") if ln.strip()]
    new_lines = [ln for ln in new_text.replace("\r", "").split("\n") if ln.strip()]
    if old_lines and len(old_lines) == len(new_lines):
        out = raw_section
        for old, new in zip(old_lines, new_lines):
            if old not in out:
                break
            out = out.replace(old, new, 1)
        else:
            return out, ""
    # Fallback: fuehrende Formatcodes behalten, Text komplett ersetzen
    prefix = ""
    m = re.match(r'^((?:&"[^"]*"|&\d+(?:\.\d+)?|&[Kk][0-9A-Fa-f]{6})+)',
                 raw_section)
    if m:
        prefix = m.group(1)
    return prefix + new_text, "Formatierung der Kopf-/Fußzeile vereinfacht"


def get_header_footer(path: str) -> dict[str, str]:
    """Liest Kopf-/Fußzeilen der Vorlage: 'HF:oddFooter.R' -> Anzeigetext."""
    out: dict[str, str] = {}
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        hf = {
            "oddHeader": ws.oddHeader, "oddFooter": ws.oddFooter,
            "evenHeader": ws.evenHeader, "evenFooter": ws.evenFooter,
            "firstHeader": ws.firstHeader, "firstFooter": ws.firstFooter,
        }
        for part, obj in hf.items():
            for section, item in (("L", obj.left), ("C", obj.center), ("R", obj.right)):
                if item.text:
                    out[f"{HF_PREFIX}{part}.{section}"] = hf_display_text(item.text)
    finally:
        wb.close()
    return out


# ---------------------------------------------------------------------------
# Layout-Pruefung (Schutz gegen verschobene/falsche Vorlagen)
# ---------------------------------------------------------------------------

LAYOUT_MIN_SIMILARITY = 0.6


def _sheet_fingerprint(ws) -> dict[str, str]:
    return {
        cell.coordinate: norm_text(cell.value)
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and len(norm_text(cell.value)) >= 4
    }


def layout_fingerprints(path: str, ignore: Optional[set] = None) -> dict[str, str]:
    """Alle Textzellen der Datei (Koordinate -> normalisierter Text).

    Mehrere Tabellenblätter werden als "Blatt!Koordinate" abgelegt, damit
    die Prüfung auch bei vertauschter Blattreihenfolge funktioniert.
    `ignore` = Zellen, die ohnehin geändert werden (Werte, keine Struktur).
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        result: dict[str, str] = {}
        for i, ws in enumerate(wb.worksheets):
            for coord, text in _sheet_fingerprint(ws).items():
                if ignore and coord in ignore and i == 0:
                    continue
                result[f"{i}!{coord}"] = text
        return result
    finally:
        wb.close()


def _split_fp(fp: dict[str, str]) -> dict[int, dict[str, str]]:
    sheets: dict[int, dict[str, str]] = {}
    for key, text in fp.items():
        idx, coord = key.split("!", 1)
        sheets.setdefault(int(idx), {})[coord] = text
    return sheets


def layout_similarity(reference: dict[str, str], other: dict[str, str]) -> float:
    """Anteil der Vorlagen-Texte, die sich in der Datei wiederfinden.

    Tolerant gegen verschobene Zeilen/Spalten (Text steht irgendwo im Blatt),
    gegen Kleinigkeiten (>= 90 % Textähnlichkeit an gleicher Stelle) und gegen
    geänderte Blattreihenfolge (bestes Blatt zählt).
    """
    if not reference:
        return 1.0
    ref = {k.split("!", 1)[1]: t for k, t in reference.items()
           if k.startswith("0!")} or {k.split("!", 1)[-1]: t
                                       for k, t in reference.items()}
    best = 0.0
    for sheet in (_split_fp(other) or {0: {}}).values():
        texts = set(sheet.values())
        matches = 0
        for coord, text in ref.items():
            here = sheet.get(coord)
            if here == text or text in texts:
                matches += 1
            elif here and difflib.SequenceMatcher(None, here, text).ratio() >= 0.9:
                matches += 1
        best = max(best, matches / len(ref))
    return best


_NUM_RE = re.compile(r"^-?\d+([.,]\d+)?$")


def _resolve_target(rule: Rule, texts: dict[str, str],
                    anchor_of: dict[str, str]) -> tuple[str, str]:
    """
    Findet die Zielzelle robust, auch wenn die Datei ein verschobenes Layout hat
    (z. B. mehr Zutaten-Zeilen als die Vorlage):
    1. ueber den Zellinhalt selbst (Textzellen)
    2. ueber die Beschriftung neben/ueber dem Feld + relativen Versatz
    3. Fallback: absolute Koordinate
    """
    base_row, base_col = coord_key(rule.coord)

    if rule.expected.strip():
        cands = [c for c, t in texts.items()
                 if norm_text(t) == norm_text(rule.expected)]
        if cands:
            cands.sort(key=lambda c: (abs(coord_key(c)[0] - base_row),
                                      abs(coord_key(c)[1] - base_col)))
            best = cands[0]
            return anchor_of.get(best, best), "inhalt"

    if rule.anchor_text.strip() and rule.anchor_coord:
        a_col_s, a_row = split_coord(rule.anchor_coord)
        a_col = col_to_idx(a_col_s)
        t_col_s, t_row = split_coord(rule.coord)
        d_col = col_to_idx(t_col_s) - a_col
        d_row = t_row - a_row
        cands = [c for c, t in texts.items()
                 if norm_text(t) == norm_text(rule.anchor_text)]
        if not cands and _row_key(rule.anchor_text):
            cands = [c for c, t in texts.items()
                     if _row_key(t) == _row_key(rule.anchor_text)]
        if cands:
            cands.sort(key=lambda c: (abs(coord_key(c)[0] - a_row),
                                      abs(coord_key(c)[1] - a_col)))
            f_col_s, f_row = split_coord(cands[0])
            target = f"{idx_to_col(col_to_idx(f_col_s) + d_col)}{f_row + d_row}"
            return anchor_of.get(target, target), "anker"

    return anchor_of.get(rule.coord, rule.coord), "koordinate"


def _apply_hf_rule(root: ET.Element, rule: Rule, verify: bool,
                   result: FileResult) -> None:
    """Wendet eine Regel auf Kopf-/Fußzeile an (z. B. 'HF:oddFooter.R')."""
    try:
        part, section = split_hf_coord(rule.coord)
    except ValueError as exc:
        result.cells.append(CellResult(
            coord=rule.coord, label=rule.label, old="", new=rule.new,
            status="fehler", message=str(exc)))
        return
    hf_el = root.find(f"{{{MAIN_NS}}}headerFooter")
    el = hf_el.find(f"{{{MAIN_NS}}}{part}") if hf_el is not None else None
    raw = (el.text or "") if el is not None else ""
    if not raw:
        result.cells.append(CellResult(
            coord=rule.coord, label=rule.label, old="", new=rule.new,
            status="uebersprungen",
            message="Kopf-/Fußzeilen-Teil nicht vorhanden"))
        return
    sections, order = _hf_split_sections(raw)
    sec_raw = sections.get(section, "")
    old_plain = hf_display_text(sec_raw)
    if verify and norm_text(old_plain) != norm_text(rule.expected):
        result.cells.append(CellResult(
            coord=rule.coord, label=rule.label, old=old_plain, new=rule.new,
            status="uebersprungen",
            message="Aktueller Text weicht vom Vorlagentext ab"))
        return
    if rule.mode == "replace" and rule.find and rule.find in sec_raw:
        new_raw, note = sec_raw.replace(rule.find, rule.new), ""
    else:
        new_raw, note = _hf_smart_replace(sec_raw, rule.new)
    sections[section] = new_raw
    el.text = _hf_join_sections(sections, order)
    result.cells.append(CellResult(
        coord=rule.coord, label=rule.label, old=old_plain, new=rule.new,
        status="geaendert", message=note))


def _looks_numeric(text: str) -> bool:
    text = text.strip()
    # fuehrende Nullen (z. B. Revision "003") muessen als Text erhalten bleiben
    return bool(_NUM_RE.match(text)) and not re.match(r"^-?0\d", text)


def _set_cell_value(cell: ET.Element, new_text: str) -> None:
    """Setzt den Zellwert, behaelt den Stil (s-Attribut) bei."""
    for child in list(cell):
        if child.tag in (f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is", f"{{{MAIN_NS}}}f"):
            cell.remove(child)
    new_text = new_text.replace("\r\n", "\n").replace("\r", "\n")
    if new_text == "":
        cell.attrib.pop("t", None)
        return
    if _looks_numeric(new_text):
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
        v.text = new_text.strip().replace(",", ".")
    else:
        cell.set("t", "inlineStr")
        is_el = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
        t_el = ET.SubElement(is_el, f"{{{MAIN_NS}}}t")
        t_el.set(f"{{{XML_NS}}}space", "preserve")
        t_el.text = new_text


def _find_or_create_cell(sheet_data: ET.Element, coord: str) -> ET.Element:
    """Sucht ein <c>-Element; legt Zeile/Zelle bei Bedarf sortiert an."""
    col_s, row_num = split_coord(coord)
    col_idx = col_to_idx(col_s)

    rows = {int(r.get("r")): r for r in sheet_data.findall(f"{{{MAIN_NS}}}row")}
    row_el = rows.get(row_num)
    if row_el is None:
        row_el = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_num)})
        inserted = False
        for i, existing in enumerate(list(sheet_data)):
            if int(existing.get("r")) > row_num:
                sheet_data.insert(i, row_el)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row_el)

    for cell in row_el.findall(f"{{{MAIN_NS}}}c"):
        if cell.get("r") == coord:
            return cell

    new_cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": coord})
    inserted = False
    for i, existing in enumerate(row_el.findall(f"{{{MAIN_NS}}}c")):
        ex_col, _ = split_coord(existing.get("r"))
        if col_to_idx(ex_col) > col_idx:
            row_el.insert(i, new_cell)
            inserted = True
            break
    if not inserted:
        row_el.append(new_cell)
    return new_cell


_REF_RE = re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)")


def _shift_rows_in_ref(ref: str, at: int, delta: int = 1) -> str:
    """Verschiebt alle Zeilennummern >= `at` in einem Bezug/Bereich um `delta`."""
    return _REF_RE.sub(
        lambda m: f"{m.group(1)}{int(m.group(2)) + (delta if int(m.group(2)) >= at else 0)}",
        ref)


def _shift_vml_rows(xml_text: str, shifts: list[tuple[int, int]]) -> str:
    """Verschiebt VML-Steuerelemente zusammen mit eingefuegten Tabellenzeilen."""
    def shift_zero_based(row: int) -> int:
        for at, delta in shifts:
            if row >= at - 1:
                row += delta
        return row

    anchor_re = re.compile(
        r"(<(?:[\w.-]+:)?Anchor\b[^>]*>)(.*?)(</(?:[\w.-]+:)?Anchor\s*>)",
        re.DOTALL)

    def shift_anchor(match: re.Match) -> str:
        parts = match.group(2).split(",")
        if len(parts) != 8:
            raise ValueError("Ungueltiger VML-Zellanker.")
        for index in (2, 6):
            row_match = re.fullmatch(r"(\s*)(-?\d+)(\s*)", parts[index])
            if row_match is None:
                raise ValueError("Ungueltige Zeile in VML-Zellanker.")
            row = shift_zero_based(int(row_match.group(2)))
            parts[index] = f"{row_match.group(1)}{row}{row_match.group(3)}"
        return match.group(1) + ",".join(parts) + match.group(3)

    xml_text = anchor_re.sub(shift_anchor, xml_text)
    row_re = re.compile(
        r"(<(?:[\w.-]+:)?Row\b[^>]*>)(\s*)(-?\d+)(\s*)(</(?:[\w.-]+:)?Row\s*>)")

    def shift_row_tag(match: re.Match) -> str:
        row = shift_zero_based(int(match.group(3)))
        return (match.group(1) + match.group(2) + str(row) + match.group(4)
                + match.group(5))

    return row_re.sub(shift_row_tag, xml_text)


def _shift_drawing_rows(xml_text: str, shifts: list[tuple[int, int]]) -> str:
    """Verschiebt DrawingML-Anker fuer Checkboxen und andere Objekte."""
    row_re = re.compile(
        r"(<(?:[\w.-]+:)?row\b[^>]*>)(\s*)(-?\d+)(\s*)(</(?:[\w.-]+:)?row\s*>)")

    def shift_row(match: re.Match) -> str:
        row = int(match.group(3))
        for at, delta in shifts:
            if row >= at - 1:
                row += delta
        return (match.group(1) + match.group(2) + str(row) + match.group(4)
                + match.group(5))

    return row_re.sub(shift_row, xml_text)


def _shift_sheet_controls(root: ET.Element,
                          shifts: list[tuple[int, int]]) -> None:
    """
    Verschiebt die Zeilenanker der Formularsteuerelemente (Kontrollkaestchen)
    im <controls>-Block des Arbeitsblatts. Excel liest die Position von dort,
    nicht aus der VML-/DrawingML-Datei.
    """
    for row_el in root.iter(f"{{{XDR_NS}}}row"):
        row = int((row_el.text or "0").strip())
        for at, delta in shifts:
            if row >= at - 1:
                row += delta
        row_el.text = str(max(row, 0))


def _merge_bottom(root: ET.Element, coord: str) -> int:
    """Letzte Zeile des verbundenen Bereichs, der `coord` enthaelt (sonst die eigene)."""
    col_s, row = split_coord(coord)
    col = col_to_idx(col_s)
    bottom = row
    mc = root.find(f"{{{MAIN_NS}}}mergeCells")
    for m in (mc.findall(f"{{{MAIN_NS}}}mergeCell") if mc is not None else []):
        ref = m.get("ref", "")
        if ":" not in ref:
            continue
        (c1, r1), (c2, r2) = (split_coord(x) for x in ref.split(":"))
        if r1 <= row <= r2 and col_to_idx(c1) <= col <= col_to_idx(c2):
            bottom = max(bottom, r2)
    return bottom


def _insert_row(root: ET.Element, sheet_data: ET.Element, after_row: int,
                model_row: int, clone_merges: bool = True,
                model_el: Optional[ET.Element] = None) -> int:
    """
    Fuegt direkt unter `after_row` eine leere Zeile ein (Stil und
    Zellverbunde von `model_row` uebernommen); alles darunter rutscht nach
    unten. Gibt die neue Zeilennummer zurueck.
    """
    at = after_row + 1
    m = lambda tag: f"{{{MAIN_NS}}}{tag}"  # noqa: E731

    for row_el in sheet_data.findall(m("row")):
        r = int(row_el.get("r"))
        if r == model_row and model_el is None:
            model_el = row_el
        if r >= at:
            row_el.set("r", str(r + 1))
            for c in row_el.findall(m("c")):
                col_s, _ = split_coord(c.get("r"))
                c.set("r", f"{col_s}{r + 1}")

    new_row = ET.Element(m("row"), {"r": str(at)})
    if model_el is not None:
        for k, v in model_el.attrib.items():
            if k != "r":
                new_row.set(k, v)
        for c in model_el.findall(m("c")):
            col_s, _ = split_coord(c.get("r"))
            nc = ET.SubElement(new_row, m("c"), {"r": f"{col_s}{at}"})
            if c.get("s"):
                nc.set("s", c.get("s"))
    children = list(sheet_data)
    pos = next((i for i, e in enumerate(children)
                if int(e.get("r")) > at), len(children))
    sheet_data.insert(pos, new_row)

    mc = root.find(m("mergeCells"))
    if mc is not None:
        clones = []
        for mcell in mc.findall(m("mergeCell")):
            ref = mcell.get("ref", "")
            if ":" in ref:
                (c1, r1), (c2, r2) = (split_coord(x) for x in ref.split(":"))
                if clone_merges and r1 == r2 == model_row:
                    clones.append(f"{c1}{at}:{c2}{at}")
            mcell.set("ref", _shift_rows_in_ref(ref, at))
        for ref in clones:
            ET.SubElement(mc, m("mergeCell"), {"ref": ref})
        mc.set("count", str(len(mc.findall(m("mergeCell")))))

    dim = root.find(m("dimension"))
    if dim is not None and dim.get("ref"):
        dim.set("ref", _shift_rows_in_ref(dim.get("ref"), at))
    for tag, attr in (("conditionalFormatting", "sqref"),
                      ("dataValidation", "sqref"), ("hyperlink", "ref")):
        for el in root.iter(m(tag)):
            if el.get(attr):
                el.set(attr, _shift_rows_in_ref(el.get(attr), at))
    for brk in root.iter(m("brk")):
        if brk.get("id") and int(brk.get("id")) >= at:
            brk.set("id", str(int(brk.get("id")) + 1))
    return at


def _shift_print_area(wb_xml: str, at: int, delta: int = 1) -> str:
    """Zieht Druckbereiche in workbook.xml bei eingefuegten/geloeschten Zeilen nach."""
    def fix(m):
        return m.group(1) + _shift_rows_in_ref(m.group(2), at, delta) + m.group(3)
    return re.sub(
        r'(<definedName[^>]*name="_xlnm\.Print_Area"[^>]*>)([^<]*)(</definedName>)',
        fix, wb_xml)


def _find_insert_position(rule: Rule, texts: dict[str, str]) -> Optional[str]:
    """Zelle in der Datei, unter deren Zeile das fehlende Feld eingefuegt wird."""
    base_row, _ = coord_key(rule.coord)
    for ctx in rule.context_above:
        cands = [c for c, t in texts.items() if norm_text(t) == norm_text(ctx)]
        if cands:
            cands.sort(key=lambda c: abs(coord_key(c)[0] - base_row))
            return cands[0]
    return None


def _mask(text: str) -> str:
    """Ziffern und x-Platzhalter egalisieren (Datum 01.02.2025 ~ xx.xx.xxxx)."""
    return re.sub(r"[0-9xX]", "#", norm_text(text).lower())


def _similar_cell(rule: Rule, texts: dict[str, str],
                  anchor_of: dict[str, str]) -> Optional[str]:
    """Aehnlichster Text in derselben Spalte (z. B. 'Revision: 002' zu '003')."""
    if not norm_text(rule.expected):
        return None
    col_s, _ = split_coord(rule.coord)
    exp = _mask(rule.expected)
    if sum(ch.isalpha() and ch != "#" for ch in exp) < 3:
        return None  # reine Zahlen/Kurzwerte sind zu mehrdeutig

    def ratio(t):
        return difflib.SequenceMatcher(None, _mask(t), exp).ratio()

    own = anchor_of.get(rule.coord, rule.coord)
    if own in texts and ratio(texts[own]) >= 0.6:
        return own
    best, best_ratio = None, 0.0
    for c, t in texts.items():
        if split_coord(c)[0] != col_s:
            continue
        r = ratio(t)
        if r > best_ratio:
            best, best_ratio = c, r
    return anchor_of.get(best, best) if best and best_ratio >= 0.75 else None


# ---------------------------------------------------------------------------
# Zeilenstruktur an die Vorlage angleichen
# ---------------------------------------------------------------------------

def _row_key(text: str) -> str:
    """Vergleichsschluessel einer Beschriftung (ohne Satzzeichen/Ziffern)."""
    return re.sub(r"[^a-zäöüß]+", "", norm_text(text).lower())


def _keyed_rows(texts: dict[str, str]) -> list[tuple[int, str]]:
    """Zeilen mit Beschriftung in Spalte A (Zeile, Schluessel), aufsteigend."""
    rows: dict[int, str] = {}
    for coord, t in texts.items():
        col, row = split_coord(coord)
        if col == "A" and _row_key(t):
            rows[row] = t
    return [(r, _row_key(rows[r])) for r in sorted(rows)]


def _delete_rows(root: ET.Element, sheet_data: ET.Element,
                 start: int, end: int) -> None:
    """Loescht die Zeilen start..end; alles darunter rutscht nach oben."""
    m = lambda tag: f"{{{MAIN_NS}}}{tag}"  # noqa: E731
    n = end - start + 1
    for row_el in list(sheet_data.findall(m("row"))):
        r = int(row_el.get("r"))
        if start <= r <= end:
            sheet_data.remove(row_el)
        elif r > end:
            row_el.set("r", str(r - n))
            for c in row_el.findall(m("c")):
                col_s, _ = split_coord(c.get("r"))
                c.set("r", f"{col_s}{r - n}")
    mc = root.find(m("mergeCells"))
    if mc is not None:
        for mcell in list(mc.findall(m("mergeCell"))):
            ref = mcell.get("ref", "")
            if ":" in ref:
                (_, r1), (_, r2) = (split_coord(x) for x in ref.split(":"))
                if start <= r1 and r2 <= end:
                    mc.remove(mcell)
                    continue
            mcell.set("ref", _shift_rows_in_ref(ref, end + 1, -n))
        mc.set("count", str(len(mc.findall(m("mergeCell")))))
    dim = root.find(m("dimension"))
    if dim is not None and dim.get("ref"):
        dim.set("ref", _shift_rows_in_ref(dim.get("ref"), end + 1, -n))
    for brk in root.iter(m("brk")):
        if brk.get("id") and int(brk.get("id")) > end:
            brk.set("id", str(int(brk.get("id")) - n))


def _style_sig(cell) -> tuple:
    f, b, a = cell.font, cell.border, cell.alignment
    col = lambda c: (c.rgb if c is not None and c.type == "rgb" else  # noqa: E731
                     (c.theme, c.tint) if c is not None and c.type == "theme" else None)
    return (f.name, f.sz, f.b, f.i, f.u, col(f.color),
            cell.fill.patternType, col(cell.fill.fgColor),
            b.left.style, b.right.style, b.top.style, b.bottom.style,
            a.horizontal, a.vertical, bool(a.wrap_text), cell.number_format)


def _sheet_sigs(path: str, only_styled: bool) -> dict:
    """Stilsignatur -> s-Index (Datei) bzw. Koordinate -> Signatur (Vorlage)."""
    wb = openpyxl.load_workbook(path)
    try:
        ws = wb.worksheets[0]
        with zipfile.ZipFile(path) as z:
            raw = ET.fromstring(z.read(_first_sheet_path(z)))
        raw_s = {c.get("r"): c.get("s") for c in raw.iter(f"{{{MAIN_NS}}}c")
                 if c.get("s")}
        out = {}
        for row in ws.iter_rows():
            for c in row:
                if c.coordinate in raw_s:
                    sig = _style_sig(c)
                    if only_styled:
                        out.setdefault(sig, raw_s[c.coordinate])
                    else:
                        out[c.coordinate] = sig
        return out
    finally:
        wb.close()


def sync_structure(template_path: str, root: ET.Element,
                   sheet_data: ET.Element, shared: list[str],
                   shifts: list[tuple[int, int]],
                   source_path: Optional[str] = None) -> list[str]:
    """
    Gleicht die Zeilen der Spezifikation an die Vorlage an: Zeilen, die nur in
    der Vorlage stehen, werden eingefuegt, abweichende Zeilen getauscht und
    Zeilen, die in der Vorlage fehlen, entfernt. Verglichen wird die
    Beschriftung in Spalte A; Produktdaten in passenden Zeilen bleiben
    unangetastet. Gibt eine Liste von Meldungen zurueck.
    """
    wb = openpyxl.load_workbook(template_path, data_only=True)
    try:
        tws = wb.worksheets[0]
        tcells = {}
        for row in tws.iter_rows():
            for c in row:
                if c.value is not None and norm_text(c.value):
                    tcells[c.coordinate] = c.value
        t_merges = [(r.min_col, r.min_row, r.max_col, r.max_row)
                    for r in tws.merged_cells.ranges]
        t_heights = {r: d.height for r, d in tws.row_dimensions.items()
                     if d.height}
    finally:
        wb.close()

    tkeys = _keyed_rows({c: str(v) for c, v in tcells.items()})
    cells = {c.get("r"): c for c in sheet_data.iter(f"{{{MAIN_NS}}}c")}
    stexts = {r: t for r, c in cells.items()
              if (t := _cell_text(c, shared)) and norm_text(t)}
    skeys = _keyed_rows(stexts)

    def block_end(keys, i, texts_rows_max, bottom_fn):
        """Letzte Zeile des Blocks von keys[i] (bis zur naechsten Beschriftung)."""
        row = keys[i][0]
        last = bottom_fn(row)
        if i + 1 < len(keys):
            return max(row, min(last, keys[i + 1][0] - 1))
        return last

    t_bottom_by_row: dict[int, int] = {}
    for c1, r1, c2, r2 in t_merges:
        for r in range(r1, r2 + 1):
            t_bottom_by_row[r] = max(t_bottom_by_row.get(r, r), r2)

    def t_bottom(row):
        return t_bottom_by_row.get(row, row)

    def s_bottom(row):
        return _merge_bottom(root, f"A{row}")

    ops = difflib.SequenceMatcher(
        None, [k for _, k in skeys], [k for _, k in tkeys],
        autojunk=False).get_opcodes()
    notes: list[str] = []
    t_sigs: Optional[dict] = None
    sig_to_s: dict = {}
    # von unten nach oben, damit die Zeilennummern oberhalb gueltig bleiben
    for tag, i1, i2, j1, j2 in reversed(ops):
        if tag == "equal":
            continue
        t_start = tkeys[j1][0] if j1 < j2 else None
        t_end = (block_end(tkeys, j2 - 1, 0, t_bottom) if j1 < j2 else None)
        if i1 < i2:
            s_start = skeys[i1][0]
            s_end = block_end(skeys, i2 - 1, 0, s_bottom)
        elif i1 < len(skeys):
            s_start, s_end = skeys[i1][0], None
        else:
            s_start = (s_bottom(skeys[-1][0]) + 1) if skeys else 1
            s_end = None
        labels = [stexts[f"A{skeys[i][0]}"] for i in range(i1, i2)]
        rows_now = {int(r.get("r")): r
                    for r in sheet_data.findall(f"{{{MAIN_NS}}}row")}
        model_src = rows_now.get(s_start)
        if model_src is None:
            model_src = rows_now.get(s_start - 1)
        model_el = copy.deepcopy(model_src) if model_src is not None else None
        if s_end is not None:
            _delete_rows(root, sheet_data, s_start, s_end)
            shifts.append((s_end + 1, -(s_end - s_start + 1)))
        if t_start is not None:
            n = t_end - t_start + 1
            for k in range(n):
                _insert_row(root, sheet_data, s_start - 1 + k, 0,
                            clone_merges=False, model_el=model_el)
                shifts.append((s_start + k, 1))
            mc = root.find(f"{{{MAIN_NS}}}mergeCells")
            for c1, r1, c2, r2 in t_merges:
                if t_start <= r1 and r2 <= t_end and mc is not None:
                    ref = (f"{idx_to_col(c1)}{r1 - t_start + s_start}:"
                           f"{idx_to_col(c2)}{r2 - t_start + s_start}")
                    ET.SubElement(mc, f"{{{MAIN_NS}}}mergeCell", {"ref": ref})
            if mc is not None:
                mc.set("count", str(len(mc.findall(f"{{{MAIN_NS}}}mergeCell"))))
            rows_by_num = {int(r.get("r")): r
                           for r in sheet_data.findall(f"{{{MAIN_NS}}}row")}
            if t_sigs is None and source_path:
                t_sigs = _sheet_sigs(template_path, False)
                sig_to_s = _sheet_sigs(source_path, True)
            for coord, sig in (t_sigs or {}).items():
                col_s, r = split_coord(coord)
                s_id = sig_to_s.get(sig)
                if t_start <= r <= t_end and s_id:
                    cell = _find_or_create_cell(
                        sheet_data, f"{col_s}{r - t_start + s_start}")
                    cell.set("s", s_id)
            for coord, value in tcells.items():
                col_s, r = split_coord(coord)
                if t_start <= r <= t_end:
                    cell = _find_or_create_cell(
                        sheet_data, f"{col_s}{r - t_start + s_start}")
                    _set_cell_value(cell, _plain(value))
            for r in range(t_start, t_end + 1):
                row_el = rows_by_num.get(r - t_start + s_start)
                if row_el is not None and t_heights.get(r):
                    row_el.set("ht", str(t_heights[r]))
                    row_el.set("customHeight", "1")
            new_labels = [str(tcells[f"A{tkeys[j][0]}"]).strip() for j in range(j1, j2)]
        else:
            new_labels = []
        if tag == "insert":
            notes.append("Zeile(n) eingefügt: " + "; ".join(new_labels))
        elif tag == "delete":
            notes.append("Zeile(n) entfernt: " + "; ".join(norm_text(x) for x in labels))
        else:
            notes.append("Zeile(n) getauscht: " + "; ".join(norm_text(x) for x in labels)
                         + " → " + "; ".join(new_labels))
    notes.reverse()
    return notes


_VERSION_RE = re.compile(
    r"revision|version|\brev\b|stand|index|datum|erstellt|ersetzt|supersede|date",
    re.IGNORECASE)
_AFTER_ROWBREAKS = ("colBreaks", "customProperties", "cellWatches",
                    "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
                    "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
                    "controls", "webPublishItems", "tableParts", "extLst")


def _is_template_text(old: str, new: str) -> bool:
    """Gleicher Festtext mit anderem Stand (Revision, 'nicht vakuumiert' -> 'vakuumiert')."""
    if norm_text(old) == norm_text(new) or not norm_text(old):
        return False
    if _NUM_RE.match(norm_text(new)):
        return False
    if difflib.SequenceMatcher(None, _mask(old), _mask(new)).ratio() < 0.6:
        return False
    letters = lambda s: re.sub(r"[^a-zäöüß]+", "", s.lower())  # noqa: E731
    return letters(old) != letters(new) or bool(_VERSION_RE.search(new))


def sync_values_and_layout(template_path: str, root: ET.Element,
                           sheet_data: ET.Element, shared: list[str]) -> list[str]:
    """
    Uebernimmt feste Texte (Revision, Vakuum-Hinweise, Kopf-/Fusszeile) aus der
    Vorlage in die Spezifikation, ohne Produktdaten zu ueberschreiben, und
    gleicht Seitenumbrueche sowie Zeilenhoehen gleicher Bloecke an.
    """
    m = lambda tag: f"{{{MAIN_NS}}}{tag}"  # noqa: E731
    wb = openpyxl.load_workbook(template_path, data_only=True)
    try:
        tws = wb.worksheets[0]
        tcells = {c.coordinate: c.value for row in tws.iter_rows() for c in row
                  if c.value is not None and norm_text(c.value)}
        t_heights = {r: d.height for r, d in tws.row_dimensions.items() if d.height}
        t_breaks = [(int(b.id), b.max, b.man) for b in
                    (tws.row_breaks.brk if tws.row_breaks else []) if b.id]
        t_hf = {}
        for part in HF_PARTS:
            obj = getattr(tws, part)
            for sec, item in (("L", obj.left), ("C", obj.center), ("R", obj.right)):
                if item.text:
                    t_hf[(part, sec)] = hf_display_text(item.text)
    finally:
        wb.close()

    cells = {c.get("r"): c for c in sheet_data.iter(m("c"))}
    stexts = {r: t for r, c in cells.items()
              if (t := _cell_text(c, shared)) and norm_text(t)}
    tkeys = _keyed_rows({c: str(v) for c, v in tcells.items()})
    skeys = _keyed_rows(stexts)
    sm = difflib.SequenceMatcher(None, [k for _, k in tkeys],
                                 [k for _, k in skeys], autojunk=False)
    pairs = [(tkeys[i + k][0], skeys[j + k][0])
             for i, j, n in sm.get_matching_blocks() for k in range(n)]
    notes: list[str] = []

    def map_row(r: int) -> int:
        off = 0
        for t_r, s_r in pairs:
            if t_r <= r:
                off = s_r - t_r
            else:
                break
        return r + off

    changed = 0
    for coord, value in tcells.items():
        col_s, r = split_coord(coord)
        target = f"{col_s}{map_row(r)}"
        old = stexts.get(target)
        if old is not None and _is_template_text(old, _plain(value)):
            _set_cell_value(cells[target], _plain(value))
            changed += 1
    if changed:
        notes.append(f"{changed} feste(r) Text(e) aus der Vorlage übernommen")

    hf_el = root.find(m("headerFooter"))
    for (part, sec), tdisp in t_hf.items():
        el = hf_el.find(m(part)) if hf_el is not None else None
        if el is None or not el.text:
            continue
        sections, order = _hf_split_sections(el.text)
        old_plain = hf_display_text(sections.get(sec, ""))
        if _is_template_text(old_plain, tdisp):
            sections[sec], _ = _hf_smart_replace(sections[sec], tdisp)
            el.text = _hf_join_sections(sections, order)
            notes.append(f"Kopf-/Fußzeile angepasst ({part}.{sec})")

    rows_by_num = {int(r.get("r")): r for r in sheet_data.findall(m("row"))}
    for (t_a, s_a), nxt in zip(pairs, pairs[1:] + [None]):
        if nxt is None:
            break
        if nxt[0] - t_a != nxt[1] - s_a:
            continue
        for k in range(nxt[0] - t_a):
            row_el, ht = rows_by_num.get(s_a + k), t_heights.get(t_a + k)
            if row_el is not None and ht:
                row_el.set("ht", str(ht))
                row_el.set("customHeight", "1")

    if t_breaks:
        rb = root.find(m("rowBreaks"))
        if rb is None:
            rb = ET.Element(m("rowBreaks"))
            kids = list(root)
            pos = next((i for i, e in enumerate(kids)
                        if e.tag.split("}")[-1] in _AFTER_ROWBREAKS), len(kids))
            root.insert(pos, rb)
        for b in list(rb):
            rb.remove(b)
        for bid, bmax, man in t_breaks:
            attrs = {"id": str(map_row(bid)), "man": "1" if man else "0"}
            if bmax:
                attrs["max"] = str(bmax)
            ET.SubElement(rb, m("brk"), attrs)
        rb.set("count", str(len(t_breaks)))
        rb.set("manualBreakCount", str(sum(1 for b in t_breaks if b[2])))
    return notes


def _plain(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def apply_to_file(
    source: str,
    target: str,
    rules: Iterable[Rule],
    verify: bool = False,
    template: Optional[str] = None,
) -> FileResult:
    """
    Wendet Ersetzungsregeln auf eine Spezifikation an und schreibt das
    Ergebnis nach `target`. Nicht betroffene ZIP-Bestandteile bleiben
    byte-identisch erhalten; VML-Steuerelemente werden bei Zeilenaenderungen
    passend mitverschoben.
    """
    result = FileResult(source=source, target=target)
    result.language = detect_language(source)
    rules = list(rules)
    if not rules and not template:
        result.error = "Keine Ersetzungsregeln ausgewaehlt."
        return result

    try:
        with zipfile.ZipFile(source) as zin:
            shared = _read_shared_strings(zin)
            sheet_path = _first_sheet_path(zin)
            sheet_xml = zin.read(sheet_path).decode("utf-8")

            ns_decls = _register_namespaces(sheet_xml)
            root = ET.fromstring(sheet_xml.encode("utf-8"))
            anchor_of = _merge_anchor_map(root)
            sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
            if sheet_data is None:
                raise ValueError("sheetData nicht gefunden.")

            cells_by_ref: dict[str, ET.Element] = {}
            texts: dict[str, str] = {}
            shifts: list[tuple[int, int]] = []
            if template:
                result.structure_notes = sync_structure(
                    template, root, sheet_data, shared, shifts, source)
                result.structure_notes += sync_values_and_layout(
                    template, root, sheet_data, shared)

            def refresh_state():
                nonlocal anchor_of
                anchor_of = _merge_anchor_map(root)
                cells_by_ref.clear()
                cells_by_ref.update(
                    {c.get("r"): c for c in sheet_data.iter(f"{{{MAIN_NS}}}c")})
                texts.clear()
                texts.update({
                    ref: t for ref, c in cells_by_ref.items()
                    if (t := _cell_text(c, shared)) and norm_text(t)})

            refresh_state()

            via_texts = {
                "inhalt": "per Zellinhalt gefunden",
                "anker": "über Beschriftung gefunden",
                "koordinate": "über Koordinate",
                "aehnlich": "über ähnlichen Text gefunden",
                "eingefuegt": "fehlendes Feld neu eingefügt",
            }

            # Von oben nach unten: Vorlagen-Koordinaten gelten, sobald alle
            # darueberliegenden Zeilen eingefuegt sind (Datei == Vorlage).
            rules.sort(key=lambda r: (r.coord.startswith(HF_PREFIX),
                                      0 if r.coord.startswith(HF_PREFIX)
                                      else coord_key(r.coord)))
            for rule in rules:
                if rule.coord.startswith(HF_PREFIX):
                    _apply_hf_rule(root, rule, verify, result)
                    continue

                coord, via = _resolve_target(rule, texts, anchor_of)
                if via == "koordinate" and rule.allow_insert:
                    similar = _similar_cell(rule, texts, anchor_of)
                    pos = _find_insert_position(rule, texts)
                    if similar:
                        coord, via = similar, "aehnlich"
                    elif pos is None:
                        result.cells.append(CellResult(
                            coord=rule.coord, label=rule.label, old="",
                            new=rule.new, status="uebersprungen",
                            message="Feld fehlt und Einfügeposition "
                                    "nicht ermittelbar"))
                        continue
                    else:
                        after = _merge_bottom(root, pos)
                        at = _insert_row(root, sheet_data, after,
                                         coord_key(pos)[0])
                        shifts.append((at, 1))
                        t_col = idx_to_col(col_to_idx(split_coord(rule.coord)[0]))
                        if (rule.anchor_text and rule.anchor_coord
                                and split_coord(rule.anchor_coord)[1]
                                == split_coord(rule.coord)[1]
                                and split_coord(rule.anchor_coord)[0] != t_col):
                            refresh_state()
                            lab = _find_or_create_cell(
                                sheet_data, f"{split_coord(rule.anchor_coord)[0]}{at}")
                            _set_cell_value(lab, rule.anchor_text)
                        refresh_state()
                        coord = anchor_of.get(f"{t_col}{at}", f"{t_col}{at}")
                        via = "eingefuegt"
                cell = cells_by_ref.get(coord)
                if cell is None:
                    cell = _find_or_create_cell(sheet_data, coord)
                    cells_by_ref[coord] = cell
                old = _cell_text(cell, shared)

                if verify and norm_text(old) != norm_text(rule.expected):
                    result.cells.append(CellResult(
                        coord=coord, label=rule.label, old=old, new=rule.new,
                        status="uebersprungen",
                        message=("Aktueller Wert weicht vom Vorlagenwert ab"
                                 if rule.expected else "Zielzelle ist nicht leer"),
                    ))
                    continue

                if rule.mode == "replace" and rule.find not in old \
                        and rule.line_start >= 0:
                    # Zeile mit anderem Datum/Kuerzel: ueber die Position ersetzen
                    lines = old.replace("\r\n", "\n").split("\n")
                    end = rule.line_start + rule.line_count
                    if end > len(lines):
                        result.cells.append(CellResult(
                            coord=coord, label=rule.label, old=old, new=rule.new,
                            status="uebersprungen",
                            message="Zeilenaufbau weicht von der Vorlage ab"))
                        continue
                    lines[rule.line_start:end] = rule.new.split("\n")
                    _set_cell_value(cell, "\n".join(lines))
                elif rule.mode == "replace":
                    if not rule.find or rule.find not in old:
                        result.cells.append(CellResult(
                            coord=coord, label=rule.label, old=old, new=rule.new,
                            status="uebersprungen",
                            message=f"Suchtext nicht gefunden: {rule.find!r}",
                        ))
                        continue
                    _set_cell_value(cell, old.replace(rule.find, rule.new))
                else:
                    _set_cell_value(cell, rule.new)

                result.cells.append(CellResult(
                    coord=coord, label=rule.label, old=old, new=rule.new,
                    status="geaendert", message=via_texts.get(via, ""),
                ))

            if shifts:
                _shift_sheet_controls(root, shifts)
            new_xml = _XML_DECL + ET.tostring(root, encoding="unicode")
            new_xml = _restore_namespaces(new_xml, ns_decls)
            _validate_sheet_xml(new_xml)

            os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
            tmp_target = target + ".tmp"
            vml_paths = (_sheet_related_paths(zin, sheet_path, "vmlDrawing")
                         if shifts else set())
            drawing_paths = (_sheet_related_paths(zin, sheet_path, "drawing")
                             if shifts else set())
            with zipfile.ZipFile(tmp_target, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == sheet_path:
                        zout.writestr(item, new_xml.encode("utf-8"))
                    elif item.filename == "xl/workbook.xml" and shifts:
                        wb_xml = zin.read(item.filename).decode("utf-8")
                        for at, delta in shifts:
                            wb_xml = _shift_print_area(wb_xml, at, delta)
                        zout.writestr(item, wb_xml.encode("utf-8"))
                    elif item.filename in vml_paths:
                        vml_xml = zin.read(item.filename).decode("utf-8")
                        vml_xml = _shift_vml_rows(vml_xml, shifts)
                        zout.writestr(item, vml_xml.encode("utf-8"))
                    elif item.filename in drawing_paths:
                        drawing_xml = zin.read(item.filename).decode("utf-8")
                        drawing_xml = _shift_drawing_rows(drawing_xml, shifts)
                        zout.writestr(item, drawing_xml.encode("utf-8"))
                    else:
                        zout.writestr(item, zin.read(item.filename))
            shutil.move(tmp_target, target)

        result.ok = True
        return result
    except Exception as exc:  # noqa: BLE001 - Fehler wird im Report gemeldet
        result.error = str(exc)
        return result


def unique_target_path(out_dir: str, source_name: str, suffix: str = "") -> str:
    """Erzeugt einen kollisionsfreien Zielpfad."""
    stem, ext = os.path.splitext(source_name)
    candidate = os.path.join(out_dir, f"{stem}{suffix}{ext}")
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(out_dir, f"{stem}{suffix} ({counter}){ext}")
        counter += 1
    return candidate


def process_batch(
    sources: Iterable[str],
    rules: Iterable[Rule],
    out_dir: str,
    suffix: str = "",
    verify: bool = False,
    progress: Optional[Callable[[int, int, FileResult], None]] = None,
    template: Optional[str] = None,
) -> list[FileResult]:
    """Verarbeitet mehrere Spezifikationen nacheinander in einen Zielordner."""
    sources = list(sources)
    results: list[FileResult] = []
    for i, src in enumerate(sources, start=1):
        target = unique_target_path(out_dir, os.path.basename(src), suffix)
        res = apply_to_file(src, target, rules, verify=verify, template=template)
        results.append(res)
        if progress:
            progress(i, len(sources), res)
    return results
