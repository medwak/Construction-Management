@echo off
REM Builds a standalone WorkforceManager.exe (no Python needed to run it afterwards).
REM Put this file in the same folder as construction_workers.py and double-click it.

cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
python -m PyInstaller --onefile --windowed --name "WorkforceManager" construction_workers.py

echo.
echo Done! Your app is here: dist\WorkforceManager.exe
echo (The database file workforce.db will be created next to the .exe on first run.)
pause
