$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

$appUrl = "http://127.0.0.1:8501"
$venvDir = Join-Path $PSScriptRoot ".venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$requirementsPath = Join-Path $PSScriptRoot "requirements.txt"
$launcherPath = Join-Path $PSScriptRoot "launcher.py"
$requirementsHashPath = Join-Path $venvDir ".requirements.sha256"
$pythonDownloadUrl = "https://www.python.org/downloads/windows/"
$launchLogPath = Join-Path $PSScriptRoot "ProCardApp-launch.log"

function Test-AppRunning {
    param([string]$Url)

    try {
        Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Find-PythonCommand {
    $versionCheck = 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'
    $candidates = @(
        @{ Command = "py"; Args = @("-3") },
        @{ Command = "python"; Args = @() }
    )

    foreach ($candidate in $candidates) {
        try {
            & $candidate.Command @($candidate.Args + @("-c", $versionCheck)) | Out-Null
            if ($LASTEXITCODE -eq 0) {
                return $candidate
            }
        }
        catch {
        }
    }

    return $null
}

function Get-RequirementsHash {
    param([string]$Path)

    return (Get-FileHash -Path $Path -Algorithm SHA256).Hash
}

function Show-PythonInstallPrompt {
    param([string]$DownloadUrl)

    Write-Host ""
    Write-Host "Python 3.10 or newer is required to run ProCard Reconciliation App." -ForegroundColor Yellow
    Write-Host "Download the 64-bit Windows installer from:" -ForegroundColor Yellow
    Write-Host $DownloadUrl -ForegroundColor Cyan
    Write-Host ""

    $response = Read-Host "Open the Python download page now? (Y/N)"
    if ($response -match '^(?i)y(?:es)?$') {
        Start-Process $DownloadUrl | Out-Null
    }
}

try {
    Start-Transcript -Path $launchLogPath -Append | Out-Null

    if (Test-AppRunning -Url $appUrl) {
        Write-Host "ProCard Reconciliation App is already running." -ForegroundColor Green
        Start-Process $appUrl | Out-Null
        exit 0
    }

    if (-not (Test-Path $requirementsPath)) {
        throw "requirements.txt was not found in $PSScriptRoot"
    }

    if (-not (Test-Path $launcherPath)) {
        throw "launcher.py was not found in $PSScriptRoot"
    }

    $pythonCommand = Find-PythonCommand
    if ($null -eq $pythonCommand) {
        Show-PythonInstallPrompt -DownloadUrl $pythonDownloadUrl
        throw "Python 3.10 or newer is required. Install it, then run launch_app.bat again."
    }

    if (-not (Test-Path $venvPython)) {
        Write-Host "Creating local Python environment..." -ForegroundColor Yellow
        & $pythonCommand.Command @($pythonCommand.Args + @("-m", "venv", $venvDir))
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create the local Python environment."
        }
    }

    $expectedHash = Get-RequirementsHash -Path $requirementsPath
    $installedHash = ""
    if (Test-Path $requirementsHashPath) {
        $installedHash = (Get-Content $requirementsHashPath -Raw).Trim()
    }

    if ($installedHash -ne $expectedHash) {
        Write-Host "Installing app dependencies (first run may take a few minutes)..." -ForegroundColor Yellow
        & $venvPython -m pip install --upgrade pip
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to upgrade pip in the local environment."
        }

        & $venvPython -m pip install -r $requirementsPath
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to install app dependencies. Check your internet connection and try again."
        }

        Set-Content -Path $requirementsHashPath -Value $expectedHash -NoNewline
    }

    Write-Host "Verifying installed packages..." -ForegroundColor Yellow
    $verifyOutput = & $venvPython -c "import streamlit, pandas, pdfplumber, openpyxl; print('All packages OK')" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "One or more required packages failed to import:`n$($verifyOutput -join "`n")`n`nTry deleting the .venv folder and running launch_app.bat again."
    }
    Write-Host ($verifyOutput -join "`n") -ForegroundColor Green

    Write-Host "Starting ProCard Reconciliation App..." -ForegroundColor Green
    & $venvPython $launcherPath
    $appExitCode = $LASTEXITCODE
    if ($appExitCode -ne 0) {
        throw "The app exited with error code $appExitCode. Review the output above for the Python error, or check: $launchLogPath"
    }
}
catch {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "ProCard App failed to start." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host ""
    if (Test-Path $launchLogPath) {
        Write-Host "--- Last 30 lines of $launchLogPath ---" -ForegroundColor Yellow
        Get-Content $launchLogPath -Tail 30 | ForEach-Object { Write-Host $_ }
        Write-Host "--- End of log ---" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Full log saved to: $launchLogPath" -ForegroundColor Cyan
    exit 1
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {
    }
}