@echo off
setlocal
chcp 65001 >nul
set "RB_PACKAGE_SCRIPT=%~dp0app\tools\package-release.ps1"
set "RB_POWERSHELL=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%RB_PACKAGE_SCRIPT%" (
  echo Package script was not found:
  echo %RB_PACKAGE_SCRIPT%
  pause
  exit /b 1
)
if not exist "%RB_POWERSHELL%" set "RB_POWERSHELL=powershell.exe"

"%RB_POWERSHELL%" -NoProfile -ExecutionPolicy Bypass -File "%RB_PACKAGE_SCRIPT%" -SharedFolderOnly
set "RB_EXIT=%ERRORLEVEL%"
if not "%RB_EXIT%"=="0" (
  echo.
  echo Failed to create the shared-folder release.
) else (
  echo.
  echo The shared-folder release was created successfully.
)
pause
exit /b %RB_EXIT%
