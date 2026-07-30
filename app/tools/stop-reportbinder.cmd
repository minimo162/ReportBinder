@echo off
set "SCRIPT=%~dp0stop-reportbinder.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
pause
