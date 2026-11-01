@echo off
setlocal
cd /d "%~dp0"
net session >nul 2>&1
if %errorlevel%==0 goto :run
powershell.exe -NoProfile -Command "try { Start-Process -FilePath '%~f0' -Verb RunAs -ErrorAction Stop; exit 0 } catch { exit 1 }"
if %errorlevel%==0 exit /b
echo.
echo Administrator approval is required to install the scheduler.
echo.
pause
exit /b 1
:run
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\Install.ps1"
endlocal
