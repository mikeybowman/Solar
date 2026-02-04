@echo off
title Solar - SHA Key Generator

echo.
echo ================================================================
echo   Solar - SHA Key Generator for GitHub Releases
echo ================================================================
echo.

if not exist "dist\Solar.exe" (
    echo ERROR: Solar.exe not found in dist folder!
    echo Please build the executable first using build_exe.bat
    echo.
    pause
    exit /b 1
)

echo Generating SHA keys for Solar.exe...
echo.

:: Generate MD5, SHA1, and SHA256 hashes
echo MD5:
certutil -hashfile "dist\Solar.exe" MD5 | findstr /v "hash"

echo.
echo SHA1:
certutil -hashfile "dist\Solar.exe" SHA1 | findstr /v "hash"

echo.
echo SHA256:
certutil -hashfile "dist\Solar.exe" SHA256 | findstr /v "hash"

echo.
echo ================================================================

:: Save hashes to a file for easy copying
echo Solar GDD Builder v1.0 - File Hashes > "dist\CHECKSUMS.txt"
echo Generated on %DATE% at %TIME% >> "dist\CHECKSUMS.txt"
echo. >> "dist\CHECKSUMS.txt"

echo MD5: >> "dist\CHECKSUMS.txt"
certutil -hashfile "dist\Solar.exe" MD5 | findstr /v "hash" >> "dist\CHECKSUMS.txt"

echo. >> "dist\CHECKSUMS.txt"
echo SHA1: >> "dist\CHECKSUMS.txt"
certutil -hashfile "dist\Solar.exe" SHA1 | findstr /v "hash" >> "dist\CHECKSUMS.txt"

echo. >> "dist\CHECKSUMS.txt"
echo SHA256: >> "dist\CHECKSUMS.txt"
certutil -hashfile "dist\Solar.exe" SHA256 | findstr /v "hash" >> "dist\CHECKSUMS.txt"

echo.
echo Checksums saved to: dist\CHECKSUMS.txt
echo.
echo For GitHub releases, use the SHA256 hash above.
echo The checksums file has been created for easy reference.
echo.

:: Show file information
echo FILE INFORMATION:
echo ================
for %%A in (dist\Solar.exe) do (
    echo File: Solar.exe
    echo Size: %%~zA bytes
    echo Date: %%~tA
)

echo.
pause
