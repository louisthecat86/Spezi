"""
SpeziSmith - Kern-Engine

Liest Spezifikations-Vorlagen (xlsx), stellt jedes Feld als waehlbare
Ersetzungsregel bereit und schreibt geaenderte Werte per ZIP/XML-Surgery
zurueck. Dadurch bleiben eingebettete Bilder (Firmenlogo), VML-Grafiken,
ActiveX-Steuerelemente, Formatierungen und Druckeinstellungen vollstaendig
erhalten - im Gegensatz zu einem openpyxl-Roundtrip.
"""

from __future__ import annotations

import os
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

def layout_fingerprints(path: str) -> dict[str, str]:
    """Alle Textzellen (Koordinate -> normalisierter Text) einer Datei."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        return {
            cell.coordinate: norm_text(cell.value)
            for row in ws.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and len(norm_text(cell.value)) >= 4
        }
    finally:
        wb.close()


def layout_similarity(reference: dict[str, str], other: dict[str, str]) -> float:
    """Anteil der Vorlagen-Textzellen, die in der Datei identisch stehen."""
    if not reference:
        return 1.0
    matches = sum(1 for coord, text in reference.items()
                  if other.get(coord) == text)
    return matches / len(reference)


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
    return bool(_NUM_RE.match(text.strip()))


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


def apply_to_file(
    source: str,
    target: str,
    rules: Iterable[Rule],
    verify: bool = False,
) -> FileResult:
    """
    Wendet Ersetzungsregeln auf eine Spezifikation an und schreibt das
    Ergebnis nach `target`. Alle anderen ZIP-Bestandteile (Logo, VML,
    ActiveX, Stile, Druckeinstellungen) bleiben byte-identisch erhalten.
    """
    result = FileResult(source=source, target=target)
    result.language = detect_language(source)
    rules = list(rules)
    if not rules:
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

            cells_by_ref = {
                c.get("r"): c for c in sheet_data.iter(f"{{{MAIN_NS}}}c")
            }
            texts = {
                ref: t for ref, c in cells_by_ref.items()
                if (t := _cell_text(c, shared)) and norm_text(t)
            }

            via_texts = {
                "inhalt": "per Zellinhalt gefunden",
                "anker": "über Beschriftung gefunden",
                "koordinate": "über Koordinate",
            }

            for rule in rules:
                if rule.coord.startswith(HF_PREFIX):
                    _apply_hf_rule(root, rule, verify, result)
                    continue

                coord, via = _resolve_target(rule, texts, anchor_of)
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

                if rule.mode == "replace":
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

            new_xml = _XML_DECL + ET.tostring(root, encoding="unicode")
            new_xml = _restore_namespaces(new_xml, ns_decls)
            _validate_sheet_xml(new_xml)

            os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
            tmp_target = target + ".tmp"
            with zipfile.ZipFile(tmp_target, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == sheet_path:
                        zout.writestr(item, new_xml.encode("utf-8"))
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
) -> list[FileResult]:
    """Verarbeitet mehrere Spezifikationen nacheinander in einen Zielordner."""
    sources = list(sources)
    results: list[FileResult] = []
    for i, src in enumerate(sources, start=1):
        target = unique_target_path(out_dir, os.path.basename(src), suffix)
        res = apply_to_file(src, target, rules, verify=verify)
        results.append(res)
        if progress:
            progress(i, len(sources), res)
    return results
