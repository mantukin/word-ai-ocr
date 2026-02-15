@echo off
setlocal
title Installing Word AI OCR...

:: --- Elevation Check ---
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [*] Requesting Administrator privileges...
    powershell -Command "Start-Process -FilePath '%0' -Verb RunAs"
    exit /b
)

echo ========================================================
echo      Word AI OCR Plugin Installer (Public Folder)
echo ========================================================
echo.

:: 1. Close Word
taskkill /f /im winword.exe >nul 2>&1

:: 2. Setup Install Directory
set "INSTALL_DIR=C:\Users\Public\Documents\WordAIOCR"
set "MANIFEST_PATH=%INSTALL_DIR%\manifest.xml"
set "MANIFEST_URL=https://mantukin.github.io/word-ai-ocr/manifest-release.xml"

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: 2a. Set Permissions and Share the Folder
echo [*] Setting folder permissions...
:: Grant read permissions to Everyone (S-1-1-0 is the universal SID for Everyone)
icacls "%INSTALL_DIR%" /grant *S-1-1-0:(OI)(CI)R /T >nul 2>&1

echo [*] Configuring network share...
:: Remove old share if exists to ensure clean state
net share WordAIOCR /delete >nul 2>&1

:: Create share using PowerShell with localized "Everyone" group name
powershell -Command "$everyone = ([Security.Principal.SecurityIdentifier]'S-1-1-0').Translate([Security.Principal.NTAccount]).Value; New-SmbShare -Name 'WordAIOCR' -Path '%INSTALL_DIR%' -FullAccess $everyone -Description 'Word AI OCR Manifest Share'" >nul 2>&1

if %errorLevel% neq 0 (
    :: Fallback: Try with SID directly in net share if possible (some environments)
    :: Or try common names
    net share WordAIOCR="%INSTALL_DIR%" /grant:Все,READ >nul 2>&1
    if %errorLevel% neq 0 (
        net share WordAIOCR="%INSTALL_DIR%" /grant:Everyone,READ >nul 2>&1
    )
)

if %errorLevel% == 0 (
    echo [V] Folder shared as: \\%COMPUTERNAME%\WordAIOCR
) else (
    echo [!] WARNING: Failed to create network share automatically.
    echo [!] Please share '%INSTALL_DIR%' manually with 'Read' permissions for 'Everyone'.
)

:: 3. Download Manifest
echo [*] Downloading manifest from GitHub...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%MANIFEST_URL%' -OutFile '%MANIFEST_PATH%'"

if not exist "%MANIFEST_PATH%" (
    echo [!] ERROR: Download failed. Check internet connection.
    pause
    exit /b
)

:: 4. Verify Content (Check if it's XML, not 404 HTML)
findstr /i "<!DOCTYPE html>" "%MANIFEST_PATH%" >nul
if %errorlevel% equ 0 (
    echo [!] ERROR: The manifest URL returned a 404 Page. 
    echo [!] Please wait a few minutes for GitHub Pages to update and try again.
    del "%MANIFEST_PATH%"
    pause
    exit /b
)

:: 5. Clear Cache (Aggressive)
echo [*] Clearing Office cache...
rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
rmdir /s /q "%userprofile%\AppData\Local\Microsoft\Office\16.0\Wef" >nul 2>&1
del /f /q "%userprofile%\AppData\Local\Microsoft\Office\16.0\Wef\*" >nul 2>&1

:: 6. Clean up old Developer Registry Key (Fixes broken links)
echo [*] Cleaning up old registry keys...
set "REG_KEY_DEV=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"
set "ADDIN_ID=9bb7a975-b568-4b2d-9683-39fd2900118f"

:: Remove the specific debugging value and subkey if present
reg delete "%REG_KEY_DEV%" /v "%ADDIN_ID%" /f >nul 2>&1
reg delete "%REG_KEY_DEV%\%ADDIN_ID%" /f >nul 2>&1

:: Also clean up from 'Manifests' subkey just in case
reg delete "%REG_KEY_DEV%\Manifests" /v "WordAIOCR" /f >nul 2>&1

:: 7. Register as Trusted Shared Folder Catalog (PowerShell Method)
echo [*] Adding to Trusted Catalogs via PowerShell...
powershell -Command "$regPath = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{97560B04-394E-4545-985A-61266052E2A6}'; $url = '\\%COMPUTERNAME%\WordAIOCR'; if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }; Set-ItemProperty -Path $regPath -Name 'Id' -Value '{97560B04-394E-4545-985A-61266052E2A6}'; Set-ItemProperty -Path $regPath -Name 'Url' -Value $url; Set-ItemProperty -Path $regPath -Name 'Flags' -Value 1 -Type DWord;"

if %errorLevel% == 0 (
    echo [V] Registry updated successfully.
) else (
    echo [!] ERROR: Failed to update registry.
)

echo.
echo ========================================================
echo [V] SUCCESS! Plugin registered and Catalog trusted.
echo ========================================================
echo.
echo 1. Open Word.
echo 2. Go to: Insert -> My Add-ins -> Shared Folder (SHARED FOLDER).
echo 3. You should see 'Word AI OCR' there.
echo.
pause
