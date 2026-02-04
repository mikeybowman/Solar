@echo off
title Solar - One-Click Launcher

echo.
echo ================================================================
echo   Solar - Open Source GDD Builder
echo   One-Click Launcher
echo ================================================================
echo.

:: Check if we have the built executable first
if exist "Solar.exe" (
    echo Starting Solar executable...
    start "" "Solar.exe"
    exit
)

if exist "dist\Solar.exe" (
    echo Starting Solar executable from dist folder...
    start "" "dist\Solar.exe"
    exit
)

:: If no exe exists, try to run from Python source
echo No executable found. Attempting to run from Python source...
echo.

:: Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo To run Solar, you need either:
    echo 1. Solar.exe (download from GitHub releases)
    echo 2. Python installed from https://python.org
    echo.
    pause
    exit /b 1
)

:: Check if main file exists
if not exist "gdd_builder.py" (
    echo ERROR: gdd_builder.py not found!
    echo Make sure you're in the correct Solar directory.
    echo.
    pause
    exit /b 1
)

:: Install requirements if needed
if exist "requirements.txt" (
    echo Installing requirements...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo WARNING: Failed to install some requirements.
        echo Solar may not work correctly.
        echo.
    )
)

echo Starting Solar...
python gdd_builder.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to start Solar
    echo.
    pause
)
