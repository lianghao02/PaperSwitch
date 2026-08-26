@echo off
chcp 65001 >nul
title PaperSwitch 紙張排版工坊

set "DIST_EXE=%~dp0dist\publish\PaperSwitch.exe"
set "DEV_EXE=%~dp0dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\win-x64\PaperSwitch.exe"

if exist "%DIST_EXE%" (
    start "" "%DIST_EXE%"
    exit /b 0
)

if exist "%DEV_EXE%" (
    start "" "%DEV_EXE%"
    exit /b 0
)

echo [PaperSwitch] 尚未建置發行成品，正在進行自動建置...
powershell -ExecutionPolicy Bypass -File "%~dp0dotnet-src\scripts\build.ps1"

if exist "%DIST_EXE%" (
    start "" "%DIST_EXE%"
    exit /b 0
) else (
    echo [PaperSwitch] 建置失敗，請確認 .NET 8 SDK 環境。
    pause
    exit /b 1
)
