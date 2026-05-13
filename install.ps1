<#
.SYNOPSIS
    Installs KuroAgent Excel Add-in for development.
    Installs Node/dependencies, registers the `excel` CLI, and verifies setup.

.DESCRIPTION
    Downloads manifest from GitHub, installs npm deps, links the CLI.
    No UTF-8 special characters - ASCII only.

.EXAMPLE
    # Full install
    .\install.ps1

    # Install from local repo (skip download)
    .\install.ps1 -LocalOnly

.NOTES
    - Requires PowerShell 5.1+ (Windows)
    - Installs Node.js if not present
#>

param(
    [switch]$LocalOnly
)

$AddInName = "KuroAgent"

# --- Functions ---
function Write-Step        { Write-Host "`n>>> $($args -join ' ')" -ForegroundColor Cyan }
function Write-Success     { Write-Host "    [OK] $($args -join ' ')" -ForegroundColor Green }
function Write-Warning-C   { Write-Host "    [!!] $($args -join ' ')" -ForegroundColor Yellow }
function Write-Error-C     { Write-Host "    [!!] $($args -join ' ')" -ForegroundColor Red }

# --- Determine repo location ---
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageJsonPath = Join-Path $ScriptDir "package.json"
if (Test-Path $PackageJsonPath) {
    $RepoDir = $ScriptDir
    Write-Step "Running in existing repo: $RepoDir"
} else {
    $CloneDir = Join-Path $env:USERPROFILE ".kuroagent"
    if (-not (Test-Path $CloneDir)) {
        Write-Step "Cloning repository ..."
        git clone https://github.com/yannassoumou/excel.git $CloneDir
        Write-Success "Cloned to $CloneDir"
    } else {
        Write-Step "Repository already cloned, pulling latest ..."
        Set-Location $CloneDir
        git pull 2>$null | Out-Null
    }
    $RepoDir = $CloneDir
}

Set-Location $RepoDir

# --- Check Node.js ---
Write-Step "Checking prerequisites"

if (Get-Command node -ErrorAction SilentlyContinue) {
    $nodeVer = & node -v
    Write-Success "Node.js $nodeVer"
} else {
    Write-Step "Node.js not found. Installing ..."
    # Try winget first
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        winget install OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        Write-Success "Node.js installed: $(& node -v)"
    } else {
        # Fallback: download installer
        $InstallerUrl = "https://nodejs.org/dist/v20.20.0/node-v20.20.0-x64.msi"
        $InstallerFile = "$env:TEMP\node-installer.msi"
        Write-Step "Downloading Node.js installer ..."
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerFile -UseBasicParsing
        Write-Step "Running Node.js installer (silent) ..."
        Start-Process msiexec.exe -ArgumentList "/i `"$InstallerFile`" /qn /norestart" -Wait
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        Remove-Item $InstallerFile -Force
    }
}

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Error-C "npm not available. Aborting."
    exit 1
}

Write-Success "npm $(& npm -v)"

# --- Install dependencies ---
Write-Step "Installing npm dependencies"
& npm install --no-audit --no-fund --loglevel=error | Out-Null
Write-Success "Dependencies installed"

# --- Install dev certs ---
Write-Step "Installing dev certificates for HTTPS server"
try {
    & npx office-addin-dev-certs install --machine 2>$null | Out-Null
    Write-Success "Dev certificates installed"
} catch {
    Write-Warning-C "Dev certs step returned non-zero (may already be installed)"
}

# --- Register the `excel` CLI ---
Write-Step "Registering 'excel' CLI"

$NpmGlobalBin = & npm bin -g 2>$null
if (-not $NpmGlobalBin) {
    $NpmGlobalBin = Join-Path $env:APPDATA "npm"
}

try {
    & npm link 2>$null | Out-Null
    Write-Success "npm link -- excel CLI on PATH"
} catch {
    # Fallback: manual copy
    $BinDir = Join-Path $env:APPDATA "npm"
    if (-not (Test-Path $BinDir)) { New-Item -ItemType Directory -Path $BinDir -Force | Out-Null }
    # Create a wrapper .cmd that calls node
    $CmdPath = Join-Path $BinDir "excel.cmd"
    $NodePath = (Get-Command node).Source
    $ExcelScript = Join-Path $RepoDir "bin\excel"
    Set-Content -Path $CmdPath -Value "@`"$NodePath`" `"$ExcelScript`" %*"
    Write-Success "Created CLI wrapper at $CmdPath"
    # Ensure APPDATA\npm is on PATH
    $UserPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    if ($UserPath -notlike "*$BinDir*") {
        [System.Environment]::SetEnvironmentVariable("Path", "$UserPath;$BinDir", "User")
        $env:Path += ";$BinDir"
    }
}

# --- Verify ---
Write-Step "Verifying installation"
if (Get-Command excel -ErrorAction SilentlyContinue) {
    Write-Success "excel CLI found: $(Get-Command excel).Source"
    & excel --version
} else {
    Write-Error-C "excel CLI not on PATH. Refresh terminal or add $NpmGlobalBin to PATH."
    exit 1
}

# --- Manifest check ---
Write-Step "Checking manifest"
try {
    & npx office-addin-manifest validate manifest.xml 2>$null | Out-Null
    Write-Success "manifest.xml valid"
} catch {
    Write-Warning-C "Manifest validation returned non-zero (may be OK for dev)"
}

# --- Summary ---
Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  $AddInName installed successfully" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Quick start:" -ForegroundColor Cyan
Write-Host "    excel                  Start dev server + sideload"
Write-Host "    excel --stop           Stop dev server"
Write-Host "    excel -m path\to.xml   Use custom manifest"
Write-Host "    excel --no-open        Server only"
Write-Host "" -ForegroundColor Cyan
Write-Host "  Dev server runs on: https://localhost:3000" -ForegroundColor Cyan
Write-Host "  Repo location:        $RepoDir" -ForegroundColor Cyan
Write-Host ""
