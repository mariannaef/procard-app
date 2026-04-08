@echo off
REM ProCard Reconciliation App Launcher
REM Starts app and opens browser when localhost is ready.

setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "APP_EXE=%SCRIPT_DIR%ProCardApp\ProCardApp.exe"
set "PORT=8501"
set "URL=http://localhost:%PORT%"

powershell -Command "try { $null = Invoke-WebRequest -Uri '%URL%' -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo.
    echo ProCard Reconciliation App is already running.
    start "" "%URL%"
    exit /b 0
)

if not exist "%APP_EXE%" (
    echo.
    echo ERROR: ProCardApp.exe not found.
    echo Expected: %APP_EXE%
    echo Keep launch_app.bat and ProCardApp folder together.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting ProCard Reconciliation App...
start "" /min "%APP_EXE%"

set "WAIT_TIME=0"
set "MAX_WAIT=120"

:check_server
timeout /t 1 /nobreak >nul

powershell -Command "try { $null = Invoke-WebRequest -Uri '%URL%' -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    start "" "%URL%"
    exit /b 0
)

set /a WAIT_TIME=%WAIT_TIME%+1
if %WAIT_TIME% lss %MAX_WAIT% goto check_server

echo.
echo App is still starting or failed to start.
echo Opening browser anyway at %URL%
start "" "%URL%"
echo If page does not load, wait 30 seconds and refresh.
exit /b 0
