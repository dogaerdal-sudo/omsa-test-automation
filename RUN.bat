@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo First setup has not been completed. Running SETUP_AND_RUN.bat...
  call SETUP_AND_RUN.bat
  exit /b %errorlevel%
)
call .venv\Scripts\activate.bat
echo Starting OMSA Test Automation at http://localhost:8501
python -m streamlit run app.py
