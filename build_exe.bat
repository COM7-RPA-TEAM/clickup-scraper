@echo off
REM ==========================================================
REM  Build clickup-scraper.exe (single-file Windows executable)
REM  Output: dist\clickup-scraper.exe
REM ==========================================================
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating virtual environment...
  py -3.13 -m venv .venv || py -m venv .venv
)

echo [2/3] Installing dependencies...
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt pyinstaller

echo [3/3] Building executable...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile ^
  --name clickup-scraper ^
  --add-data "templates;templates" ^
  --collect-submodules openpyxl ^
  --hidden-import openpyxl.cell._writer ^
  app.py

echo.
echo ==========================================================
echo  Done!  ->  dist\clickup-scraper.exe
echo  Copy that single file to any Windows PC and double-click.
echo ==========================================================
pause
