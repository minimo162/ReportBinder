@echo off
set "ROOT=%~dp0..\..\"
echo Starting ReportBinder diagnostics mode: en
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%app\launch.ps1" -Mode en -Diagnostics
echo.
echo Latest log: %ROOT%app\logs\startup-en-latest.log
echo URL shortcut: %ROOT%app\logs\ReportBinder-en.url
pause
