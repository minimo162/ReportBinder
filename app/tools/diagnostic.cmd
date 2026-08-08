@echo off
set "ROOT=%~dp0..\..\"
echo Starting ReportBinder diagnostics mode
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%app\launch.ps1" -Diagnostics
echo.
echo Latest log: %%LOCALAPPDATA%%\ReportBinder\logs\startup-workspace-latest.log
echo URL shortcut: %%LOCALAPPDATA%%\ReportBinder\logs\ReportBinder-workspace.url
pause
