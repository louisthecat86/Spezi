"""Headless-Smoke-Test der GUI (unter xvfb-run ausfuehren).

Baut die komplette Oberflaeche auf, simuliert das Anklicken eines Feldes in
der Vorlagen-Ansicht, das Eintragen von Aenderungen, das Hereinziehen zweier
Dateien (Batch) und prueft das Ergebnis.
"""

import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from spezismith.gui import SpeziSmithApp  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG = os.path.join(ROOT, "spezismith_config.json")
if os.path.exists(CFG):
    os.remove(CFG)  # deterministischer Startzustand


def pump(app, seconds):
    end = time.time() + seconds
    while time.time() < end:
        app.root.update()
        time.sleep(0.02)


def main():
    app = SpeziSmithApp()
    pump(app, 1.0)

    # 1) Vorlagen-Ansicht wurde aufgebaut
    assert len(app.sheet.cells) > 100, f"zu wenige Zellen: {len(app.sheet.cells)}"
    print(f"OK: {len(app.sheet.cells)} klickbare Zellen im Raster (DE)")

    # 2) Feld A1 anklicken und Wert eintragen (wie der Anwender)
    app.sheet.select("A1")
    assert app.sel_coord_var.get() == "A1"
    app.sel_new_var.set("SmokeTest GmbH")
    pump(app, 0.2)
    assert app.sheet.cells["A1"].new_text == "SmokeTest GmbH"
    assert app.sheet.cells["A1"].changed
    rules = app._collect_rules()
    assert len(rules) == 1 and rules[0].new == "SmokeTest GmbH"
    print("OK: Regel aus Klick+Tippen erzeugt:", rules[0].coord)

    # 2b) Anker-Information: leeres Feld hinter 'Vakuum'
    assert "D78" in app.sheet.cells
    a_text, a_coord = app.sheet.anchor_of("D78")
    assert a_text == "Vakuum" and a_coord == "C78", (a_text, a_coord)
    print("OK: Anker für D78 =", a_coord, repr(a_text))

    # 2c) Kopf-/Fußzeile ist wählbar
    assert "HF:oddFooter.R" in app.sheet.cells
    app.sheet.select("HF:oddFooter.R")
    assert "Fußzeile" in app.sel_coord_var.get()
    app.sel_new_var.set("Datum 23.10.2026\n\n\nErsetzt 10.06.2026")
    pump(app, 0.2)
    rules = app._collect_rules()
    assert len(rules) == 2
    hf_rule = next(r for r in rules if r.coord == "HF:oddFooter.R")
    assert hf_rule.mode == "set"
    print("OK: Fußzeilen-Regel erzeugt:", hf_rule.coord)

    # 3) Zwei Dateien vorbereiten und "hereinziehen"
    tmp = tempfile.mkdtemp(prefix="spezismith_smoke_")
    srcs = []
    for i in range(2):
        dst = os.path.join(tmp, f"Spezi_{i}.xlsx")
        shutil.copy(os.path.join(ROOT, "templates", "vorlage_de.xlsx"), dst)
        srcs.append(dst)
    out_dir = os.path.join(tmp, "out")
    app.out_dir.set(out_dir)
    app.out_mode.set("folder")
    app._handle_files(srcs)

    # Worker abwarten
    for _ in range(200):
        pump(app, 0.05)
        if not (app._worker and app._worker.is_alive()):
            break
    pump(app, 0.3)  # Queue leeren lassen

    # 4) Ergebnisse pruefen
    outs = sorted(os.listdir(out_dir))
    assert len(outs) == 2, f"erwartete 2 Ausgaben, fand {outs}"
    import openpyxl
    for name in outs:
        wb = openpyxl.load_workbook(os.path.join(out_dir, name))
        val = wb.worksheets[0]["A1"].value
        foot = wb.worksheets[0].oddFooter.right.text
        wb.close()
        assert "SmokeTest GmbH" in val, f"A1 nicht geaendert in {name}: {val!r}"
        assert "23.10.2026" in foot, f"Fußzeile nicht geaendert in {name}"
    print("OK: Batch von 2 Dateien verarbeitet, A1 + Fußzeile geaendert:", outs)

    log_items = app.log.get_children()
    assert len(log_items) == 2, f"Log-Eintraege: {len(log_items)}"
    print("OK: Log enthaelt beide Dateien")

    # 4b) Layout-Pruefung: eine fremde Datei wird uebersprungen
    import openpyxl as _opx
    fremd = os.path.join(tmp, "Fremd.xlsx")
    wb = _opx.Workbook()
    wb.active["A1"] = "Ganz andere Datei ohne Spezi-Layout"
    wb.save(fremd)
    wb.close()
    app._handle_files([fremd])
    for _ in range(200):
        pump(app, 0.05)
        if not (app._worker and app._worker.is_alive()):
            break
    pump(app, 0.3)
    last = app.log.item(app.log.get_children()[-1], "values")
    assert "Layout" in last[1], last
    print("OK: Layout-Prüfung greift:", last[1][:60])

    # 5) Sprachwechsel auf Englisch (Aenderung auf A1 bleibt erhalten)
    app.lang_var.set("English")
    app._on_language_change()
    pump(app, 1.0)
    assert app.sheet.cells["A1"].new_text == "SmokeTest GmbH"
    assert app.sheet.cells["A1"].changed
    assert "Product" in app.sheet.cells["A3"].orig_text
    assert "HF:oddHeader.L" in app.sheet.cells
    print(f"OK: Sprachwechsel nach EN ({len(app.sheet.cells)} Zellen), "
          f"Aenderung auf A1 erhalten")

    # 5b) Zoom aendern (ganze Seite kleiner)
    app.lang_var.set("Deutsch")
    app._on_language_change()
    pump(app, 0.8)
    app.zoom_var.set("70 %")
    app._on_zoom_change()
    pump(app, 0.8)
    assert app.sheet.scale == 0.7
    assert len(app.sheet.cells) > 100
    print("OK: Zoom 70 % aufgebaut")
    app.zoom_var.set("100 %")
    app._on_zoom_change()
    pump(app, 0.5)

    # 6) Suche
    app.search_var.set("Artikelnummer")
    app._search_next()
    info = app.search_info_var.get()
    assert "Treffer 1/" in info, info
    first = app.sheet.selected
    app._search_next()
    assert app.sheet.selected != first or len(app.sheet._matches) == 1
    print("OK: Suche funktioniert:", info)

    # 7) Feld leeren + Zuruecknehmen
    app.sheet.select("B71")
    app.sel_clear_var.set(True)
    app._on_clear_toggled()
    assert app.sheet.cells["B71"].new_text == ""
    app._revert_current()
    assert app.sheet.cells["B71"].new_text is None
    print("OK: Feld leeren / Zuruecknehmen funktioniert")

    print("\nSMOKE-TEST ERFOLGREICH")
    app.root.destroy()
    if os.path.exists(CFG):
        os.remove(CFG)


if __name__ == "__main__":
    main()
