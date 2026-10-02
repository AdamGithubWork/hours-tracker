@echo off
title Hours Tracker

cd /d "%~dp0"

echo ==============================
echo        Hours Tracker
echo ==============================
echo.

echo Checking Python...
python --version >nul 2>&1

if errorlevel 1 (
    echo Python was not found.
    echo.
    echo Please install Python from:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b
)

echo Python found.
echo.

echo Installing required packages...
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo Failed to install the required packages.
    echo.
    pause
    exit /b
)

echo.
echo Starting Hours Tracker...
echo.

python main.py

echo.
echo ==============================
echo Tracker closed.
echo ==============================
pause