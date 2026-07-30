@echo off
set "ROOT=%~dp0..\..\"
echo Cleaning ReportBinder top folder...
del /q "%ROOT%start-ja.cmd" 2>nul
del /q "%ROOT%start-en.cmd" 2>nul
del /q "%ROOT%start-ja-diagnostic.cmd" 2>nul
del /q "%ROOT%start-en-diagnostic.cmd" 2>nul
del /q "%ROOT%stop-reportbinder.cmd" 2>nul
del /q "%ROOT%ReportBinder-ja.url" 2>nul
del /q "%ROOT%ReportBinder-en.url" 2>nul
del /q "%ROOT%ENCODING_FIX_NOTE.txt" 2>nul
del /q "%ROOT%STARTUP_FIX_NOTE.txt" 2>nul
del /q "%ROOT%STARTUP_FIX_V2_NOTE.txt" 2>nul
del /q "%ROOT%.gitignore" 2>nul
echo Done.
pause
