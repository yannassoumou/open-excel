<#
.SYNOPSIS
    Uninstalls KuroAgent Excel Add-in and development tools.

.DESCRIPTION
    Removes registry sideloading keys, unlinks the Excel CLI,
    and optionally cleans up the local installation directory.
    ASCII-only.

.PARAMETER All
    Also remove the .kuroagent installation directory.

.EXAMPLE
    .\uninstall.ps1
    .\uninstall.ps1 -All

.NOTES
    - Requires PowerShell 5.1+ (Windows)
#>

param(
    [switch]$All
)

$AddInName = "KuroAgent"
$Guid = "14254940-5dfe-46ec-b860-a8291f526990"
$RegPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\$Guid"

# --- Functions ---
function Write-Step        { Write-Host "`n>>> $($args -join ' ')" -ForegroundColor Cyan }
function Write-Success     { Write-Host "    [OK] $($args -join ' ')" -ForegroundColor Green }
function Write-Warning-C   { Write-Host "    [!!] $($args -join ' ')" -ForegroundColor Yellow }

# --- 1. Remove registry sideloading ---
Write-Step "Removing Excel registry sideloading"

if (Test-Path $RegPath) {
    Remove-Item -Path $RegPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Success "Registry key removed"
} else {
    Write-Warning-C "No registry key found (already clean)"
}

# --- 2. Unlink the Excel CLI ---
Write-Step "Unlinking excel CLI"

# Remove CLI wrapper scripts from global npm bin
$NpmGlobalBin = & npm bin -g 2>$null
if (-not $NpmGlobalBin) {
    $NpmGlobalBin = Join-Path $env:APPDATA "npm"
}

$RemovedAny = $false

foreach ($fileName in @("excel.cmd", "excel.ps1", "excel")) {
    $f = Join-Path $NpmGlobalBin $fileName
    if (Test-Path $f) {
        Remove-Item $f -Force -ErrorAction SilentlyContinue
        Write-Success "Removed: $f"
        $RemovedAny = $true
    }
}

# Remove bin directory created by npm link
$ExcelBinDir = Join-Path $NpmGlobalBin "excel"
if (Test-Path $ExcelBinDir) {
    Remove-Item $ExcelBinDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Success "Removed: $ExcelBinDir"
    $RemovedAny = $true
}

if (-not $RemovedAny) {
    Write-Warning-C "CLI wrapper not found (may not have been installed)"
}

# Clean uninstall npm link for the package
try {
    & npm unlink -g excel-custom-functions-js 2>$null | Out-Null
    Write-Success "npm unlink completed"
} catch {
    # already clean
}

# --- 3. Remove installation directory (if --All) ---
$InstallDir = Join-Path $env:USERPROFILE ".kuroagent"

if ($All) {
    Write-Step "Removing installation directory: $InstallDir"
    if (Test-Path $InstallDir) {
        Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Success "Directory removed: $InstallDir"
    } else {
        Write-Warning-C "Directory not found: $InstallDir"
    }
} else {
    Write-Warning-C "Keeping $InstallDir (use -All flag to remove)"
}

# --- Summary ---
Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  $AddInName uninstalled successfully" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "  To reinstall:" -ForegroundColor Cyan
Write-Host "    iwr https://raw.githubusercontent.com/yannassoumou/open-excel/master/install.ps1 -OutFile install.ps1"
Write-Host "    powershell -ExecutionPolicy Bypass -File .\install.ps1"
Write-Host ""
Write-Host "  Note: Restart Excel to complete the uninstall." -ForegroundColor Yellow
Write-Host ""
