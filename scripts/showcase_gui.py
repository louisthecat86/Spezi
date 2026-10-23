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
    app.root.geometry("1120x780+30+30")

    # Beispiel-Eintraege
    demo = {"A5": "123456", "A6": "Ofenkäse Schnitte 2,5 kg", "A7": "2500",
            "A1": "Musterfirma GmbH, Musterstraße 1, 12345 Musterstadt"}
    for row in app.rows:
        if row.field.coord in demo:
            row.new_value.set(demo[row.field.coord])

    app.out_dir.set(os.path.join(ROOT, "Ausgabe"))
    app.refresh_counter()

    end = time.time() + 3.5
    while time.time() < end:
        app.root.update()
        time.sleep(0.05)
    app.root.destroy()
    if os.path.exists(CFG):
        os.remove(CFG)


if __name__ == "__main__":
    main()
