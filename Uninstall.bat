@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title DDO++ Attendance Agent - Uninstall
echo.
echo  Removes Task Scheduler jobs and Start Menu shortcuts.
echo  Does NOT delete this folder, .env, Excel files, or Python.
echo.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Uninstall
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
  echo Uninstall failed with exit code %ERR%.
) else (
  echo Uninstall complete.
)
echo.
pause
exit /b %ERR%
