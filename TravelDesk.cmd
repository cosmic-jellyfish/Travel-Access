@echo off
setlocal
cd /d "%~dp0"
net session >nul 2>&1
if %errorlevel%==0 goto :run
powershell.exe -NoProfile -Command "try { Start-Process -FilePath '%~f0' -Verb RunAs -ErrorAction Stop; exit 0 } catch { exit 1 }"
if %errorlevel%==0 exit /b
echo Administrator approval was declined. Opening with limited rights.
echo Run sync now may be unavailable.
:run
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0app\TravelDesk.ps1"
endlocal
