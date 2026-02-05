@echo off
setlocal enabledelayedexpansion
title Solar - Professional Build System

echo.
echo ================================================================
echo   Solar - Open Source GDD Builder
echo   Professional Build System v1.1.0
echo   Building self-contained executable...
echo ================================================================
echo.

:: Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

:: Check for required files
if not exist "gdd_builder.py" (
    echo ERROR: gdd_builder.py not found!
    echo Make sure you're running this from the Solar directory.
    pause
    exit /b 1
)

if not exist "solar.ico" (
    echo WARNING: solar.ico not found. Building without custom icon.
    set ICON_PARAM=
) else (
    set ICON_PARAM=--icon=solar.ico
)

if not exist "version_info.txt" (
    echo WARNING: version_info.txt not found. Building without version info.
    set VERSION_PARAM=
) else (
    set VERSION_PARAM=--version-file=version_info.txt
)

echo Installing build requirements...
echo ==============================
pip install --upgrade pip
pip install pyinstaller python-docx

echo.
echo Building Solar.exe...
echo ====================
echo This process may take 1-3 minutes depending on your system.
echo.

:: Clean previous builds
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Solar.spec" del "Solar.spec"

:: Build the executable and immediately check result
pyinstaller --onefile --windowed --name "Solar" %ICON_PARAM% %VERSION_PARAM% --add-data "requirements.txt;." --add-data "sample_gdd_project.json;." --hidden-import "docx" --hidden-import "docx.shared" --hidden-import "docx.enum.text" --hidden-import "docx.oxml" --hidden-import "docx.oxml.ns" --hidden-import "docx.oxml.parser" --hidden-import "docx.parts" --hidden-import "docx.text" --hidden-import "docx.document" --clean --noconfirm gdd_builder.py

if !errorlevel! equ 0 (
    echo.
    echo ================================================================
    echo   BUILD SUCCESSFUL!
    echo ================================================================
    echo.
    echo Solar.exe created successfully in the 'dist' folder
    
    if exist "dist\Solar.exe" (
        echo File size: 
        for %%A in (dist\Solar.exe) do echo %%~zA bytes
    )
    
    echo.
    echo The executable includes:
    echo  - Complete Python runtime
    echo  - All required libraries (python-docx, tkinter)
    echo  - Sample project file
    echo  - Custom icon and version information
    echo.
    echo This .exe file can be distributed to any Windows computer
    echo and will run without requiring Python installation.
    echo.
    
    :: Copy additional files to dist folder
    if exist "solar.ico" copy "solar.ico" "dist\" >nul
    if exist "README.md" copy "README.md" "dist\" >nul
    if exist "LICENSE" copy "LICENSE" "dist\" >nul
    
    :: Create a release info file
    echo Solar GDD Builder v1.1.0 > "dist\RELEASE_INFO.txt"
    echo Built on %DATE% at %TIME% >> "dist\RELEASE_INFO.txt"
    echo. >> "dist\RELEASE_INFO.txt"
    echo This is a self-contained executable that includes: >> "dist\RELEASE_INFO.txt"
    echo - Solar GDD Builder application >> "dist\RELEASE_INFO.txt"
    echo - Python runtime environment >> "dist\RELEASE_INFO.txt"
    echo - Microsoft Word export functionality >> "dist\RELEASE_INFO.txt"
    echo - Sample project files >> "dist\RELEASE_INFO.txt"
    echo. >> "dist\RELEASE_INFO.txt"
    echo No additional software installation required! >> "dist\RELEASE_INFO.txt"
    echo. >> "dist\RELEASE_INFO.txt"
    echo For support and source code: >> "dist\RELEASE_INFO.txt"
    echo https://github.com/mikeybowman/Solar >> "dist\RELEASE_INFO.txt"
    
    echo Opening dist folder...
    explorer "dist"
    
    goto :success
) else (
    echo.
    echo ================================================================
    echo   BUILD FAILED!
    echo ================================================================
    echo.
    echo Build process encountered errors. Common solutions:
    echo.
    echo 1. Make sure you have a stable internet connection
    echo 2. Try running as Administrator
    echo 3. Temporarily disable antivirus software
    echo 4. Make sure no other Python processes are running
    echo 5. Try: pip install --upgrade pyinstaller python-docx
    echo.
    echo If problems persist, check the error messages above.
    echo.
    goto :failure
)

:success
echo.
echo Build process completed successfully!
pause
exit /b 0

:failure
echo.
echo Build process failed.
pause
exit /b 1