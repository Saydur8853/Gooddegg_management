@echo off
setlocal
title Gooddegg Inventory

cd /d "%~dp0"

echo Starting Gooddegg Inventory...

where py >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not available in PATH.
    echo Please install Python 3 and try again.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Could not create the virtual environment.
        pause
        exit /b 1
    )
)

echo Installing/checking required packages...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check xlrd openpyxl
if errorlevel 1 (
    echo Could not install required packages.
    pause
    exit /b 1
)

echo Server is running at http://127.0.0.1:8000
echo Close this window to stop the server.
".venv\Scripts\python.exe" app.py

pause
