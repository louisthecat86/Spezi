"""Tests fuer die SpeziSmith-Engine mit den echten Vorlagen."""

import os
import zipfile

import openpyxl
import pytest

from spezismith.engine import (
    Rule,
    analyze_template,
    apply_to_file,
    detect_language,
    process_batch,
    unique_target_path,
)

TPL_DE = os.path.join(os.path.dirname(__file__), "..", "templates", "vorlage_de.xlsx")
TPL_EN = os.path.join(os.path.dirname(__file__), "..", "templates", "vorlage_en.xlsx")


# ---------------------------------------------------------------------------
# Analyse & Spracherkennung
# ---------------------------------------------------------------------------

def test_detect_language():
    assert detect_language(TPL_DE) == "de"
    assert detect_language(TPL_EN) == "en"


def test_analyze_de_fields():
    fields = analyze_template(TPL_DE)
    by_coord = {f.coord: f for f in fields}

    # Kernfelder vorhanden
    assert "A1" in by_coord          # Firmenkopf
    assert "A5" in by_coord          # Artikelnummer (Beschriftung)
    assert "J1" in by_coord          # "Spezifikation gültig ab ..."
    assert "A44" in by_coord         # Allergene
    assert "A86" in by_coord         # maschinell erstellt

    assert "AULBACH GmbH" in by_coord["A1"].value
    assert by_coord["A5"].value == "Artikelnummer:"

    # leere Eingabefelder neben Beschriftungen wurden erkannt
    inputs_row5 = [f for f in fields if f.is_input and f.row == 5]
    assert inputs_row5, "Eingabefeld neben 'Artikelnummer:' fehlt"

    # keine doppelten Koordinaten
    assert len(by_coord) == len(fields)


def test_analyze_en_fields():
    fields = analyze_template(TPL_EN)
    by_coord = {f.coord: f for f in fields}
    assert by_coord["A3"].value == "Product-Specification"
    assert "B71" in by_coord                      # tare weight Wert 1.15
    assert by_coord["B71"].value.startswith("1.15")


# ---------------------------------------------------------------------------
# Ersetzen (ZIP-Surgery)
# ---------------------------------------------------------------------------

def _read_cell(path, coord):
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        return wb.worksheets[0][coord].value
    finally:
        wb.close()


def test_apply_full_cell_replacement(tmp_path):
    out = tmp_path / "out.xlsx"
    res = apply_to_file(TPL_DE, str(out), [
        Rule(coord="A1", new="Test1"),
        Rule(coord="A5", new="Artikelnummer (NEU):"),
    ])
    assert res.ok, res.error
    assert res.language == "de"
    assert res.changed_count == 2
    assert "Test1" in _read_cell(str(out), "A1")
    assert _read_cell(str(out), "A5") == "Artikelnummer (NEU):"


def test_logo_and_parts_preserved(tmp_path):
    """Logo, VML, ActiveX etc. bleiben byte-identisch erhalten."""
    out = tmp_path / "out.xlsx"
    res = apply_to_file(TPL_DE, str(out), [Rule(coord="A1", new="Test1")])
    assert res.ok, res.error

    with zipfile.ZipFile(TPL_DE) as a, zipfile.ZipFile(out) as b:
        names_a, names_b = set(a.namelist()), set(b.namelist())
        assert names_a == names_b, "ZIP-Struktur veraendert!"
        for name in names_a:
            if name == "xl/worksheets/sheet1.xml":
                continue
            assert a.read(name) == b.read(name), f"Part geaendert: {name}"
        # Logo explizit pruefen
        assert a.read("xl/media/image1.jpeg") == b.read("xl/media/image1.jpeg")


def test_merges_and_styles_survive(tmp_path):
    out = tmp_path / "out.xlsx"
    apply_to_file(TPL_DE, str(out), [Rule(coord="A1", new="Test1")])
    wb = openpyxl.load_workbook(out)
    try:
        ws = wb.worksheets[0]
        assert len(ws.merged_cells.ranges) >= 140
        assert ws["A1"].font is not None  # Stilinformation vorhanden
    finally:
        wb.close()


def test_replace_substring_mode(tmp_path):
    out = tmp_path / "out.xlsx"
    rule = Rule(coord="J1", new="01.01.2027", mode="replace", find="xx.xx.xxxx")
    res = apply_to_file(TPL_DE, str(out), [rule])
    assert res.ok, res.error
    assert res.changed_count == 1
    assert "01.01.2027" in _read_cell(str(out), "J1")
    assert "xx.xx.xxxx" not in str(_read_cell(str(out), "J1"))


def test_replace_substring_not_found(tmp_path):
    out = tmp_path / "out.xlsx"
    rule = Rule(coord="J1", new="X", mode="replace", find="GIBTS-NICHT")
    res = apply_to_file(TPL_DE, str(out), [rule])
    assert res.ok
    assert res.cells[0].status == "uebersprungen"


def test_verify_mismatch_skips(tmp_path):
    out = tmp_path / "out.xlsx"
    rule = Rule(coord="A1", new="Test1", expected="völlig anderer Wert")
    res = apply_to_file(TPL_DE, str(out), [rule], verify=True)
    assert res.ok
    assert res.cells[0].status == "uebersprungen"
    assert "AULBACH" in _read_cell(str(out), "A1")


def test_numeric_value_stays_numeric(tmp_path):
    out = tmp_path / "out.xlsx"
    res = apply_to_file(TPL_EN, str(out), [Rule(coord="B71", new="2,30")])
    assert res.ok, res.error
    value = _read_cell(str(out), "B71")
    assert value == 2.3  # als Zahl gespeichert, Komma wurde konvertiert


def test_write_into_empty_cell(tmp_path):
    """Leere Eingabefelder (z. B. Wert neben 'Artikelnummer:') fuellen."""
    fields = analyze_template(TPL_DE)
    target_field = next(f for f in fields if f.is_input and f.row == 5)
    out = tmp_path / "out.xlsx"
    res = apply_to_file(TPL_DE, str(out), [
        Rule(coord=target_field.coord, new="ART-12345"),
    ])
    assert res.ok, res.error
    assert _read_cell(str(out), target_field.coord) == "ART-12345"


def test_clear_cell(tmp_path):
    out = tmp_path / "out.xlsx"
    res = apply_to_file(TPL_EN, str(out), [Rule(coord="B71", new="")])
    assert res.ok, res.error
    assert _read_cell(str(out), "B71") in (None, "")


# ---------------------------------------------------------------------------
# Batch & Zielpfade
# ---------------------------------------------------------------------------

def test_unique_target_path(tmp_path):
    first = unique_target_path(str(tmp_path), "Spezi.xlsx", "_neu")
    assert first.endswith("Spezi_neu.xlsx")
    open(first, "wb").close()
    second = unique_target_path(str(tmp_path), "Spezi.xlsx", "_neu")
    assert "(1)" in second


def test_process_batch_five_files(tmp_path):
    """Fuenf Spezifikationen gleichzeitig verarbeiten."""
    sources = []
    for i in range(5):
        src = tmp_path / f"Spezi_{i}.xlsx"
        with open(TPL_DE, "rb") as fsrc, open(src, "wb") as fdst:
            fdst.write(fsrc.read())
        sources.append(str(src))

    out_dir = tmp_path / "out"
    events = []
    results = process_batch(
        sources,
        [Rule(coord="A1", new="Sammel-Test")],
        str(out_dir),
        suffix="_neu",
        progress=lambda i, n, r: events.append((i, n)),
    )
    assert len(results) == 5
    assert all(r.ok for r in results)
    assert events == [(1, 5), (2, 5), (3, 5), (4, 5), (5, 5)]
    for r in results:
        assert "Sammel-Test" in _read_cell(r.target, "A1")


def test_invalid_file_reports_error(tmp_path):
    bad = tmp_path / "kaputt.xlsx"
    bad.write_text("kein Excel")
    res = apply_to_file(str(bad), str(tmp_path / "o.xlsx"), [Rule(coord="A1", new="X")])
    assert not res.ok
    assert res.error
