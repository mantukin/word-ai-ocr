@echo off
title AI Word OCR Plugin Runner
echo Starting AI Word OCR Plugin...

:: 1. Build the project if dist is missing
if not exist "dist" (
    echo Building project...
    call npm run build
)

:: 2. Pre-configure loopback (with auto-yes)
echo Configuring localhost loopback...
echo y | call npx office-addin-dev-settings appcontainer manifest.xml --loopback

:: 3. Start the local server in background
echo Starting local server...
:: Using 'call' to ensure node starts correctly from within bat
start /b node local-server.js

:: Wait for server to be fully ready
echo Waiting for server to initialize...
timeout /t 5 /nobreak > nul

:: 4. Start Word and Sideload (with auto-yes for the second prompt)
echo Sideloading manifest and opening Word...
echo y | call npx office-addin-debugging start manifest.xml --no-debug

echo.
echo ====================================================
echo PLUGIN IS ACTIVE!
echo ====================================================
echo 1. Keep this window open.
echo 2. Go to Word -> Home Tab -> Show Task Pane.
echo 3. If you see 'File not found', wait 5 seconds and 
echo    right-click in the task pane -> Refresh.
echo.
echo To stop everything, close this window.
pause
