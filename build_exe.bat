@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found.
  pause
  exit /b 1
)

python -m pip install --upgrade pyinstaller
python -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name VBA_Certificate_Generator ^
  create_vba_cert_gui.py

echo.
echo Build finished.
echo EXE path: dist\VBA_Certificate_Generator.exe
pause
