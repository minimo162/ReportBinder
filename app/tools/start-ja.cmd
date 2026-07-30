@echo off
set "ROOT=%~dp0..\..\"
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%ROOT%app\launch.ps1" -Mode ja
