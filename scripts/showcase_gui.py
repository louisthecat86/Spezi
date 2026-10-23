"""Startet die GUI kurz mit Beispieldaten (fuer Screenshots unter xvfb)."""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from spezismith.gui import SpeziSmithApp  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG = os.path.join(ROOT, "spezismith_config.json")
if os.path.exists(CFG):
    os.remove(CFG)


def main():
    app = SpeziSmithApp()
    app.root.geometry("1240x800+20+20")

    # warten, bis die Vorlagen-Ansicht aufgebaut ist
    end = time.time() + 4.0
    while time.time() < end and not app.sheet.cells:
        app.root.update()
        time.sleep(0.05)

    # Beispiel-Aenderungen, als haette der Anwender sie eingetippt
    demo = {"A5": "123456", "A6": "Ofenkäse-Schnitte 2,5 kg", "A7": "2500",
            "J1": "Spezifikation gültig ab 01.01.2027",
            "HF:oddFooter.R": "Datum 23.10.2026\n\n\nErsetzt 10.06.2026"}
    for coord, text in demo.items():
        if coord in app.sheet.cells:
            app.sheet.set_new_value(coord, text)
    app.sheet.select("A6")
    app.sel_new_var.set("Ofenkäse-Schnitte 2,5 kg")
    app._refresh_counter()
    app.out_dir.set(os.path.join(ROOT, "Ausgabe"))

    end = time.time() + 3.0
    while time.time() < end:
        app.root.update()
        time.sleep(0.05)
    app.root.destroy()
    if os.path.exists(CFG):
        os.remove(CFG)


if __name__ == "__main__":
    main()
