@echo off
setlocal EnableExtensions
title PaperSwitch

set "DIST_EXE=%~dp0dist\publish\PaperSwitch.exe"
set "DEV_EXE=%~dp0dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\win-x64\PaperSwitch.exe"
set "BUILD_SCRIPT=%~dp0dotnet-src\scripts\build.ps1"

if exist "%DIST_EXE%" (
    start "" "%DIST_EXE%"
    exit /b 0
)

if exist "%DEV_EXE%" (
    start "" "%DEV_EXE%"
    exit /b 0
)

set "PS_HOST=pwsh.exe"
where.exe pwsh.exe >nul 2>&1
if errorlevel 1 set "PS_HOST=powershell.exe"

"%PS_HOST%" -NoProfile -ExecutionPolicy Bypass -File "%BUILD_SCRIPT%"
if errorlevel 1 goto :build_failed

if exist "%DIST_EXE%" (
    start "" "%DIST_EXE%"
    exit /b 0
)

:build_failed
echo [PaperSwitch] Build failed. Check the .NET 8 SDK and build output.
pause
exit /b 1
