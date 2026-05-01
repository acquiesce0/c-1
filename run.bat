@echo off
cd /d "%~dp0"
python app.py
if errorlevel 1 (
  echo.
  echo ============================================================
  echo The app did not start. Most common reasons:
  echo   1. Python is not installed, or "Add to PATH" was not ticked
  echo      during install. Reinstall from https://www.python.org/downloads/
  echo      and TICK the box that says "Add python.exe to PATH".
  echo   2. The openpyxl library is not installed. Open Command Prompt
  echo      and run:   pip install openpyxl
  echo ============================================================
  pause
)
