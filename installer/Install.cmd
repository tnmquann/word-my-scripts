@echo off
rem Double-click to install the MY_SCRIPTS Word add-in for the current user.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
echo.
pause
