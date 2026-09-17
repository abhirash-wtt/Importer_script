@echo off
setlocal
cd /d "%~dp0"
echo.
echo DDO++ Attendance Agent - one-click install
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
  echo Install failed with exit code %ERR%.
) else (
  echo Done. See messages above for next steps.
)
echo.
pause
exit /b %ERR%
