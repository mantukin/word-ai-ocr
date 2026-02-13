@echo off
setlocal
title Installing AI Word OCR Plugin...

echo ========================================================
echo      AI Word OCR Plugin Installer (Windows)
echo ========================================================
echo.

:: 1. Define installation directory (User's local AppData)
set "INSTALL_DIR=%APPDATA%\AIWordOCR"
set "MANIFEST_URL=https://mantukin.github.io/word-ai-ocr/manifest-release.xml"

echo [*] Creating installation folder...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: 2. Download the latest manifest
echo [*] Downloading latest manifest from GitHub...
powershell -Command "Invoke-WebRequest -Uri '%MANIFEST_URL%' -OutFile '%INSTALL_DIR%\manifest.xml'"

if %errorlevel% neq 0 (
    echo [!] Error downloading manifest. Check your internet connection.
    pause
    exit /b
)

:: 3. Add Registry Key for Trusted Catalog
echo [*] Configuring Microsoft Word settings...

:: Generate a unique ID for the catalog entry
set "CATALOG_ID={5C187B03-6C74-4228-9369-637373737373}"
set "REG_PATH=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%"

:: Add the path to the registry
reg add "%REG_PATH%" /v "Url" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
:: Set Flags to 1 (Show in Menu)
reg add "%REG_PATH%" /v "Flags" /t REG_DWORD /d 1 /f >nul
:: Set Id
reg add "%REG_PATH%" /v "Id" /t REG_SZ /d "%CATALOG_ID%" /f >nul

echo.
echo ========================================================
echo [V] SUCCESS! Installation complete.
echo ========================================================
echo.
echo 1. Open Microsoft Word.
echo 2. Go to Insert -> My Add-ins.
echo 3. Click on 'SHARED FOLDER' (Shared Folder).
echo 4. Select 'Word AI OCR' and click Add.
echo.
echo (You may need to restart Word if it was open)
echo.
pause
