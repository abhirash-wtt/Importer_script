@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title DDO++ Attendance Agent - Setup
color 0B
echo.
echo  ============================================================
echo   DDO++ Attendance Agent
echo   Office Install Package
echo  ============================================================
echo.
echo  This setup will:
echo    - Install Python 3 if missing (via winget)
echo    - Install required Python packages
echo    - Create folders and .env
echo    - Create the attendance drop folder
echo    - Register Task Scheduler: eSSL 13:00, importer 13:10 (Mon-Fri)
echo    - Add Start Menu shortcuts
echo.
echo  Right-click -^> Run as administrator if Task Scheduler fails.
echo.
pause

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
set ERR=%ERRORLEVEL%

echo.
if %ERR% neq 0 (
  color 0C
  echo  Setup FAILED with exit code %ERR%.
  echo  See messages above, then fix and run Setup.bat again.
) else (
  color 0A
  echo  Setup finished successfully.
)
echo.
pause
exit /b %ERR%
