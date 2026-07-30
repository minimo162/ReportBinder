@echo off
set "ROOT=%~dp0..\..\"
echo Starting ReportBinder diagnostics mode: ja
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%app\launch.ps1" -Mode ja -Diagnostics
echo.
echo Latest log: %ROOT%app\logs\startup-ja-latest.log
echo URL shortcut: %ROOT%app\logs\ReportBinder-ja.url
pause
