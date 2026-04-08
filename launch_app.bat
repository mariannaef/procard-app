@echo off
REM ProCard Reconciliation App Launcher
REM Creates a local environment on first run, installs requirements, and starts the app.

setlocal

set "SCRIPT_DIR=%~dp0"
set "BOOTSTRAP_PS1=%SCRIPT_DIR%bootstrap_portable.ps1"

if not exist "%BOOTSTRAP_PS1%" (
    echo.
    echo ERROR: bootstrap_portable.ps1 not found.
    echo Expected: %BOOTSTRAP_PS1%
    echo Keep the launcher files together after extracting the zip.
    echo.
    pause
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%BOOTSTRAP_PS1%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo ============================================================
    echo  ProCard App could not be started.
    echo  See ProCardApp-launch.log in this folder for details.
    echo ============================================================
)

pause
exit /b %EXIT_CODE%
