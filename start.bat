@echo off
title PaperSwitch Launcher

set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

echo ============================================================
echo   PaperSwitch - Document to PDF Converter
echo ============================================================
echo.

cd /d "%~dp0"

set PYTHON_CMD=python

if exist "..\.venv\Scripts\python.exe" (
    set PYTHON_CMD=..\.venv\Scripts\python.exe
    goto :FOUND_PYTHON
)

if exist ".venv\Scripts\python.exe" (
    set PYTHON_CMD=.venv\Scripts\python.exe
    goto :FOUND_PYTHON
)

:FOUND_PYTHON
echo [INFO] Using Python: %PYTHON_CMD%
echo.

%PYTHON_CMD% -c "import win32com, PIL, dotenv, pypdf" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing required dependencies...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Dependency installation failed.
        pause
        exit /b 1
    )
)

echo [INFO] Launching PaperSwitch Web Server...
echo.
%PYTHON_CMD% app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Server stopped with errors.
    pause
)
