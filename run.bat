@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found.
  echo Please install Python 3.9 or later, or run the packaged EXE version.
  pause
  exit /b 1
)

python create_vba_cert_gui.py
pause
