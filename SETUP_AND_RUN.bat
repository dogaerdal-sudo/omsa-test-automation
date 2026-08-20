@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo OMSA Test Automation - First Setup
echo ============================================================

set "PY_CMD="
py -3.12 --version >nul 2>&1 && set "PY_CMD=py -3.12"
if not defined PY_CMD py -3.13 --version >nul 2>&1 && set "PY_CMD=py -3.13"
if not defined PY_CMD python --version >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD py --version >nul 2>&1 && set "PY_CMD=py"

if not defined PY_CMD (
  echo ERROR: Python was not found.
  echo Install Python 3.12 or 3.13 and run this file again.
  pause
  exit /b 1
)

echo Using: %PY_CMD%

if not exist ".venv\Scripts\python.exe" (
  echo Creating local virtual environment...
  %PY_CMD% -m venv .venv
  if errorlevel 1 goto :error
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo Starting OMSA Test Automation...
echo Browser URL: http://localhost:8501
echo Keep this window open while using the application.
echo.
python -m streamlit run app.py
exit /b 0

:error
echo.
echo SETUP FAILED. Take a screenshot of the error above.
pause
exit /b 1
