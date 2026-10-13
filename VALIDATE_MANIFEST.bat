@echo off
setlocal
cd /d "%~dp0"
if not exist node_modules call npm install
call npm run validate-manifest
pause
