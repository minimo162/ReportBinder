@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-thirdparty.ps1" %*
set RC=%ERRORLEVEL%
echo.
if not "%RC%"=="0" (
  echo install-thirdparty failed. ErrorLevel=%RC%
) else (
  echo install-thirdparty completed.
)
echo.
pause
exit /b %RC%
