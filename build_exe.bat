@echo off
REM ============================================================
REM  SpeziSmith - Windows-Build (portable EXE via PyInstaller)
REM  Aufruf: Doppelklick auf build_exe.bat  oder  build_exe.bat
REM ============================================================
setlocal

echo [1/3] Abhaengigkeiten installieren ...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fehler

echo [2/3] EXE bauen ...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name SpeziSmith ^
  --icon assets\SpeziSmith.ico ^
  --add-data "templates;templates" ^
  --add-data "assets\logo.png;assets" ^
  --add-data "assets\icon.png;assets" ^
  --collect-all tkinterdnd2 ^
  --collect-all sv_ttk ^
  main.py
if errorlevel 1 goto fehler

echo [3/3] Fertig!
echo.
echo  Die portable EXE liegt hier:  dist\SpeziSmith.exe
echo  Einfach auf Stick/Festplatte kopieren und per Doppelklick starten.
echo.
pause
exit /b 0

:fehler
echo.
echo  BUILD FEHLGESCHLAGEN - siehe Meldungen oben.
pause
exit /b 1
