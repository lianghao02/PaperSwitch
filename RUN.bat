@echo off
setlocal EnableExtensions
title PaperSwitch

set "PS_HOST=pwsh.exe"
where.exe pwsh.exe >nul 2>&1
if errorlevel 1 set "PS_HOST=powershell.exe"

"%PS_HOST%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0dotnet-src\scripts\run.ps1"
if not errorlevel 1 exit /b 0

echo [PaperSwitch] Startup failed. Review the message above.
pause
exit /b 1
