@echo off
rem Double-click to remove the MY_SCRIPTS Word add-in for the current user.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
echo.
pause
