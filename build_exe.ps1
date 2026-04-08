$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

$distRootPath = Join-Path $PSScriptRoot "dist"
$portablePath = Join-Path $distRootPath "ProCardApp_portable"
$zipPath = Join-Path $PSScriptRoot "ProCardApp_portable.zip"

if (Test-Path $portablePath) {
    Remove-Item $portablePath -Recurse -Force
}

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

New-Item -ItemType Directory -Path $portablePath | Out-Null

$requiredFiles = @(
    "app.py",
    "processor.py",
    "launcher.py",
    "launch_app.bat",
    "launch_app.command",
    "bootstrap_portable.ps1",
    "bootstrap_portable.sh",
    "requirements.txt",
    "logo.png"
)

foreach ($file in $requiredFiles) {
    $sourcePath = Join-Path $PSScriptRoot $file
    if (-not (Test-Path $sourcePath)) {
        throw "Required file missing: $file"
    }

    Copy-Item -Path $sourcePath -Destination (Join-Path $portablePath $file) -Force
}

Compress-Archive -Path $portablePath -DestinationPath $zipPath -CompressionLevel Optimal

Write-Host "Build complete. Share folder: dist\ProCardApp_portable" -ForegroundColor Green
Write-Host "Portable zip created: ProCardApp_portable.zip" -ForegroundColor Green
