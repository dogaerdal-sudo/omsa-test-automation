@echo off
setlocal
cd /d "%~dp0"
echo === Python ===
if exist ".venv\Scripts\python.exe" (
  .venv\Scripts\python.exe --version
) else (
  python --version
)
echo.
echo === Microsoft Excel COM ===
if exist ".venv\Scripts\python.exe" (
  .venv\Scripts\python.exe -c "import win32com.client as w; x=w.DispatchEx('Excel.Application'); print('Excel COM: OK'); print('Excel version:', x.Version); x.Quit()"
) else (
  python -c "import win32com.client as w; x=w.DispatchEx('Excel.Application'); print('Excel COM: OK'); print('Excel version:', x.Version); x.Quit()"
)
echo.
pause
