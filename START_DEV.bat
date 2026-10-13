@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required. Install a current Node.js LTS release first.
  pause
  exit /b 1
)
if not exist node_modules (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 goto :error
)
echo Starting FormaScholar at https://localhost:3000 ...
call npm run dev
exit /b %errorlevel%
:error
echo.
echo Setup failed. Review the npm error above.
pause
exit /b 1
