@echo off
REM ProCard Reconciliation App Launcher
REM This script starts the app in the background and opens the browser automatically

setlocal enabledelayedexpansion

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"
set "APP_EXE=%SCRIPT_DIR%dist\ProCardApp\ProCardApp.exe"
set "PORT=8501"
set "URL=http://localhost:%PORT%"

REM Check if the app exe exists
if not exist "%APP_EXE%" (
    echo Error: ProCardApp.exe not found at %APP_EXE%
    echo.
    echo Make sure you extracted the ProCardApp_portable.zip and this batch file is in the same folder as the "dist" folder.
    pause
    exit /b 1
)

REM Open the browser immediately if the app is already running
powershell -Command "try { $null = Invoke-WebRequest -Uri '%URL%' -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    start "" "%URL%"
    exit /b 0
)

REM Start the app in the background (minimized)
start "" /min "%APP_EXE%"

REM Wait for the Streamlit server to start (usually takes 3-5 seconds)
echo Starting ProCard Reconciliation App...
timeout /t 4 /nobreak

REM Open the browser to the Streamlit URL
echo Opening browser...
start "" "%URL%"

exit /b 0
