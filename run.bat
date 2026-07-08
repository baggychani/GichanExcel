@echo off
setlocal
cd /d "%~dp0"

where npm.cmd >nul 2>nul
if errorlevel 1 (
  echo Node.js/npm was not found.
  echo Install Node.js LTS, then run this launcher again.
  pause
  exit /b 1
)

echo Starting GichanExcel in development mode...
echo This window must stay open while the app is running.
echo.
npm.cmd run tauri dev
set EXIT_CODE=%ERRORLEVEL%
echo.
pause
exit /b %EXIT_CODE%
