@echo off
title Solar - Setup & Installation

echo.
echo ================================================================
echo   Solar - Open Source GDD Builder
echo   Setup & Installation Script
echo ================================================================
echo.

echo Checking system requirements...
echo ==============================

:: Check Python installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    echo Alternative: Download Solar.exe from GitHub releases
    echo (No Python installation required)
    echo.
    pause
    exit /b 1
)

echo ✓ Python found
python --version

:: Check for main file
if not exist "gdd_builder.py" (
    echo.
    echo ERROR: gdd_builder.py not found!
    echo Make sure you've downloaded the complete Solar package.
    echo.
    pause
    exit /b 1
)

echo ✓ Solar source files found

echo.
echo Installing Solar dependencies...
echo ===============================

:: Upgrade pip first
python -m pip install --upgrade pip

:: Install requirements
if exist "requirements.txt" (
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo WARNING: Some dependencies failed to install.
        echo Solar may not work correctly.
        echo.
        echo Try running as Administrator if you continue to have issues.
        echo.
        pause
    ) else (
        echo ✓ All dependencies installed successfully
    )
) else (
    echo Installing core dependencies...
    pip install python-docx
)

echo.
echo ================================================================
echo   Setup Complete!
echo ================================================================
echo.
echo Solar is now ready to use. You can:
echo.
echo 1. Double-click "Solar.bat" to launch Solar
echo 2. Run "python gdd_builder.py" from command line
echo 3. Build executable with "build_exe.bat"
echo.
echo For the best experience, consider building the executable
echo so you don't need Python for future use.
echo.

:: Ask if user wants to start Solar now
set /p choice="Start Solar now? (y/n): "
if /i "%choice%"=="y" (
    echo.
    echo Starting Solar...
    python gdd_builder.py
)

echo.
echo Installation complete. Enjoy using Solar!
pause

