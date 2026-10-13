@echo off
setlocal
cd /d "%~dp0"
if not exist node_modules (
  call npm install
  if errorlevel 1 goto :error
)
call npm run typecheck
if errorlevel 1 goto :error
call npm run build
if errorlevel 1 goto :error
echo.
echo Build completed. Output: dist\
pause
exit /b 0
:error
echo.
echo Build failed. Review the error above.
pause
exit /b 1
