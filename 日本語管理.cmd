@echo off
setlocal
set "RB_LAUNCHER=%~dp0app\launch.ps1"
set "RB_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%RB_LAUNCHER%" (
  echo ReportBinder launcher was not found:
  echo %RB_LAUNCHER%
  pause
  exit /b 1
)
if not exist "%RB_POWERSHELL%" set "RB_POWERSHELL=powershell.exe"

start "" /b "%RB_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%RB_LAUNCHER%" -Mode ja
if errorlevel 1 (
  echo ReportBinder could not be started.
  pause
  exit /b 1
)
exit /b 0
