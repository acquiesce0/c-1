@echo off
cd /d "%~dp0"

REM --- Silent dependency checks ---
where python >nul 2>&1
if errorlevel 1 goto no_python

python -c "import sys; sys.exit(0 if sys.version_info >= (3,8) else 1)" >nul 2>&1
if errorlevel 1 goto bad_python

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 goto no_openpyxl

python -c "import tkinter" >nul 2>&1
if errorlevel 1 goto no_tk

REM --- pywin32 is optional (used to open Excel at the right sheet on double-click).
REM     If missing, install silently in the background; the app still works without it.
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 start "" /b python -m pip install --user pywin32 >nul 2>&1

REM --- All good, launch the app windowed
REM     If it crashes, re-run with python so the traceback is visible.
where pythonw >nul 2>&1
if errorlevel 1 goto run_console
start "" pythonw app.py
goto :eof

:run_console
python app.py
if errorlevel 1 (
  echo.
  echo ============================================================
  echo The app exited with an error. See the message above.
  echo ============================================================
  pause
)
goto :eof

:no_python
echo ============================================================
echo Python was not found on PATH.
echo Install it from https://www.python.org/downloads/
echo and TICK the box "Add python.exe to PATH" during setup.
echo ============================================================
pause
goto :eof

:bad_python
echo ============================================================
echo Your Python version is too old. Python 3.8 or newer is required.
echo Install the latest from https://www.python.org/downloads/
echo ============================================================
pause
goto :eof

:no_openpyxl
echo ============================================================
echo The "openpyxl" library is not installed.
echo Installing it now...
echo ============================================================
python -m pip install --user openpyxl
if errorlevel 1 (
  echo.
  echo Automatic install failed. Open Command Prompt and run:
  echo     pip install openpyxl
  pause
  goto :eof
)
python app.py
if errorlevel 1 pause
goto :eof

:no_tk
echo ============================================================
echo Python's "tkinter" module is missing from this install.
echo Reinstall Python from https://www.python.org/downloads/
echo and make sure "tcl/tk and IDLE" is checked during setup.
echo ============================================================
pause
goto :eof
