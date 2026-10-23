"""Headless-Smoke-Test der GUI (unter xvfb-run ausfuehren).

Baut die komplette Oberflaeche auf, simuliert das Ausfuellen eines Feldes,
das Hereinziehen zweier Dateien (Batch) und prueft das Ergebnis.
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
    pump(app, 0.5)

    # 1) Felder wurden aus der Vorlage geladen
    assert len(app.rows) > 80, f"zu wenige Felder: {len(app.rows)}"
    print(f"OK: {len(app.rows)} Felder geladen (DE)")

    # 2) Feld A1 aktivieren und Wert eintragen (Auto-Haken via Entry)
    row = next(r for r in app.rows if r.field.coord == "A1")
    row.new_value.set("SmokeTest GmbH")
    assert row.checked.get(), "Auto-Haken wurde nicht gesetzt"
    rules = app._collect_rules()
    assert len(rules) == 1 and rules[0].new == "SmokeTest GmbH"
    print("OK: Regel erzeugt:", rules[0])

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

    # 5) Sprachwechsel auf Englisch (Eingaben bleiben fuer gleiche Zellen erhalten)
    app.lang_var.set("English")
    app._on_language_change()
    pump(app, 0.3)
    row_a1 = next(r for r in app.rows if r.field.coord == "A1")
    assert row_a1.new_value.get() == "SmokeTest GmbH"
    assert any(r.field.coord == "A3" and "Product" in r.field.value
               for r in app.rows)
    print(f"OK: Sprachwechsel nach EN, {len(app.rows)} Felder, Eingabe erhalten")

    # 6) Filter (Teilstring-Suche: "A13" trifft nur Zeilen, die A13 enthalten)
    app.filter_var.set("A13")
    app._apply_filter()
    visible = [r for r in app.rows if r.l_coord.winfo_manager()]
    assert visible and all("a13" in (r.field.coord + r.field.label
                                     + r.field.value).lower()
                           for r in visible), [r.field.coord for r in visible]
    assert any(r.field.coord == "A13" for r in visible)
    app.filter_var.set("")
    print("OK: Filter funktioniert")

    print("\nSMOKE-TEST ERFOLGREICH")
    app.root.destroy()
    if os.path.exists(CFG):
        os.remove(CFG)


if __name__ == "__main__":
    main()
