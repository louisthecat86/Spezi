# SpeziSmith 🥄📄

**Spezifikationen (FB_QM_005 Produktspezifikation / Product Specification) komfortabel
anpassen** – als portable Windows-EXE, die ohne Installation überall läuft
(USB-Stick, Festplatte, Netzlaufwerk).

![SpeziSmith Oberfläche](assets/screenshot_light.png)

## Was macht das Tool?

1. **Felder analysieren:** SpeziSmith liest die hinterlegte Vorlage (Deutsch *oder*
   Englisch – wählbar oben rechts) und listet **jedes einzelne Feld** der
   Spezifikation auf: befüllte Zellen *und* leere Eingabefelder neben
   Beschriftungen.
2. **Auswählen & neuen Wert eintragen:** Haken setzen (entsteht automatisch, sobald
   man tippt) und in das Kästchen dahinter den neuen Wert schreiben.
   Beispiel: Feld `A1` enthält „Test“ → Haken an `A1`, „Test1“ eintragen.
3. **Spezis hereinziehen:** Eine oder **mehrere Spezifikationen gleichzeitig**
   (z. B. 5 markierte Dateien) in die Drag-&-Drop-Zone ziehen. Alle aktivierten
   Felder werden ersetzt; das Ergebnis landet wahlweise
   - in einem **Zielordner** (Pfad eingeben/durchsuchen) oder
   - per **„Speichern unter“-Dialog** („Download“-Gefühl) für jede Datei.

Das Firmenlogo, alle Formatierungen, VML-/ActiveX-Elemente und Druckeinstellungen
bleiben dabei **vollständig erhalten** (die Ersetzung erfolgt direkt auf
XML-Ebene, kein verlustbehafteter Excel-Roundtrip).

## Weitere Features

- 🌍 **Sprachauswahl** Deutsch/English; abgelegte Dateien mit der „falschen“
  Sprache werden erkannt und übersprungen (Warnung im Log)
- 🔎 **Filter** über Zelle, Beschriftung und aktuellen Wert
- ☑️ Zählwerk „x von y Feldern aktiv“, „Alle/Keine“, „Nur aktive anzeigen“
- ✍️ Zwei Ersetzungsmodi: **ganze Zelle** oder **Teiltext** (z. B. Datum im Kopf)
- 🛡️ Optionale Verifikation: nur ersetzen, wenn der Ist-Wert dem Vorlagenwert
  entspricht (Abweichungen werden übersprungen und gemeldet)
- 🌙 Hell/Dunkel-Theme
- 💾 Einstellungen werden portabel in `spezismith_config.json` **neben der EXE**
  gespeichert – ideal für den USB-Stick
- 📊 Fortschrittsbalken + farbiges Ergebnis-Log (Doppelklick öffnet Zielordner)

## Die fertige EXE erhalten

### Variante A – automatisch über GitHub Actions (empfohlen)

1. Repo zu GitHub pushen
2. Tab **Actions** → Workflow *„SpeziSmith EXE bauen“* → **Run workflow**
3. Nach ~2 Minuten liegt `SpeziSmith.exe` als **Artefakt** zum Download bereit

### Variante B – lokal unter Windows bauen

Voraussetzung: [Python 3.12+](https://www.python.org/downloads/) (im Installer
„Add python.exe to PATH“ aktivieren). Dann im Projektordner:

```
build_exe.bat
```

→ fertige, in sich geschlossene `dist\SpeziSmith.exe` (~20 MB), die man überallhin
kopieren kann.

## Bedienung im Detail

| Schritt | Aktion |
|---|---|
| 1 | Sprache wählen → Felder filtern → Haken setzen + neue Werte eintragen |
| 2 | Zielordner angeben *oder* „Nach jeder Datei fragen“; Dateinamen-Zusatz (Standard `_angepasst`); optional Verifikation/Teiltext-Modus |
| 3 | `.xlsx`-Dateien in die Drop-Zone ziehen (Mehrfachauswahl!) oder über die Schaltfläche auswählen |

Eigene (neuere) Vorlagenversionen lassen sich nutzen, indem die Dateien unter
`templates\` ersetzt werden – der Aufbau wird bei jedem Start neu analysiert.

## Entwicklung

```
pip install -r requirements.txt -r requirements-dev.txt
python main.py                    # GUI starten
python -m pytest tests/ -q        # Engine-Tests (15 Tests)
xvfb-run python scripts/smoke_gui.py   # GUI-Smoke-Test (Linux)
python scripts/make_assets.py     # Logo/Icon neu erzeugen
```

### Projektstruktur

```
├── main.py                  # Einstieg / PyInstaller-Ziel
├── spezismith/
│   ├── engine.py            # Analyse, Spracherkennung, XML-sicheres Ersetzen
│   ├── gui.py               # Oberfläche (ttk + Sun-Valley + DnD)
│   └── resources.py         # Pfade, portable Konfiguration
├── templates/               # eingebettete Vorlagen (DE/EN)
├── assets/                  # Logo, Icon, Screenshots
├── tests/                   # Engine-Tests mit den echten Vorlagen
├── scripts/                 # Asset-Generator, Smoke-Tests
└── build_exe.bat            # Windows-Build
```

## Technischer Hinweis zum Logo

Das Firmenlogo (`assets/logo.jpeg`, extrahiert aus der Vorlage) liegt in jeder
erzeugten Spezifikation wieder bei: Die Engine kopiert beim Schreiben alle
ZIP-Bestandteile der Originaldatei **byte-identisch** und verändert nur die
Zellinhalte im Tabellenblatt-XML.