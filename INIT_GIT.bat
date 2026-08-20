@echo off
setlocal
cd /d "%~dp0"
if exist ".git" (
  echo Git repository already exists.
) else (
  git init
)
git status
echo.
echo Note: templates_local is intentionally ignored so large/confidential Excel templates are not pushed to GitHub.
echo When ready, run:
echo   git add .
echo   git commit -m "Initial OMSA Test Automation"
pause
