@echo off
title GST Invoice Aggregator
cd /d "%~dp0"

echo =======================================================
echo   GST INVOICE CUMULATIVE AGGREGATOR
echo =======================================================
echo.

:: Detect Node runtime
set "NODE_CMD="

if exist "%~dp0node-v20.18.0-win-x64\node.exe" (
    set "NODE_CMD=%~dp0node-v20.18.0-win-x64\node.exe"
    echo Using portable Node runtime...
) else (
    where node >nul 2>nul
    if %errorlevel% equ 0 (
        set "NODE_CMD=node"
        echo Using installed Node runtime...
    )
)

if "%NODE_CMD%"=="" (
    echo Error: Node.js was not found.
    echo Please make sure the folder is fully unzipped, or install Node.js from https://nodejs.org
    echo.
    pause
    exit /b 1
)

echo Starting GST Invoice Aggregator on http://localhost:3000 ...
echo Your web browser will open in a moment.
echo (Keep this window open while using the app)
echo.

"%NODE_CMD%" server.js

pause
