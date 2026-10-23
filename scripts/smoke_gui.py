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
        wb.close()
        assert "SmokeTest GmbH" in val, f"A1 nicht geaendert in {name}: {val!r}"
    print("OK: Batch von 2 Dateien verarbeitet, A1 geaendert:", outs)

    log_items = app.log.get_children()
    assert len(log_items) == 2, f"Log-Eintraege: {len(log_items)}"
    print("OK: Log enthaelt beide Dateien")

    # 5) Sprachwechsel auf Englisch (Aenderung auf A1 bleibt erhalten)
    app.lang_var.set("English")
    app._on_language_change()
    pump(app, 1.0)
    assert app.sheet.cells["A1"].new_text == "SmokeTest GmbH"
    assert app.sheet.cells["A1"].changed
    assert "Product" in app.sheet.cells["A3"].orig_text
    print(f"OK: Sprachwechsel nach EN ({len(app.sheet.cells)} Zellen), "
          f"Aenderung auf A1 erhalten")

    # 6) Suche
    app.search_var.set("Article")
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
