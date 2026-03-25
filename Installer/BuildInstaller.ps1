# Professional Build Script for Riyan Revit Tools Installer
# This script will download Inno Setup if missing and compile the installer.

$issPath = ".\RiyanSetup.iss"
$outputExe = "..\RiyanSetup_v1.0.8.exe"

Write-Host "--- Riyan Revit Tools Installer Builder ---" -ForegroundColor Cyan

# 1. Check for Inno Setup
$iscc = Get-Command iscc -ErrorAction SilentlyContinue
if (-not $iscc) {
    $localPath = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    if (Test-Path $localPath) {
        $iscc = [PSCustomObject]@{ Source = $localPath }
    } else {
        Write-Host "Inno Setup not found. Attempting to install via winget..." -ForegroundColor Yellow
        winget install -e --id JRSoftware.InnoSetup
        $iscc = Get-Command iscc -ErrorAction SilentlyContinue
    }
}

if (-not $iscc) {
    Write-Host "Error: Inno Setup is still not found. Please install it manually from https://jrsoftware.org/isdl.php" -ForegroundColor Red
    exit 1
}

# 2. Compile the installer
Write-Host "Compiling installer..." -ForegroundColor Green
& $iscc.Source $issPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "Success! Your professional installer is ready at: $(Resolve-Path $outputExe)" -ForegroundColor Green
} else {
    Write-Host "Error: Compilation failed." -ForegroundColor Red
}
