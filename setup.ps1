<#
.SYNOPSIS
    Unified setup for kuroagent Excel Add-in.
    Install, update, or uninstall in one script.

.DESCRIPTION
    All setup operations in a single entry point.
    ASCII-only safe for PowerShell download.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\setup.ps1          # Interactive menu
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 install   # Install
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 update    # Update
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 uninstall  # Uninstall (keep .kuroagent)
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 purge     # Uninstall + remove .kuroagent

.NOTES
    - Requires PowerShell 5.1+ (Windows)
#>

param(
    [Parameter(Position = 0)]
    [ValidateSet("install", "update", "uninstall", "purge", "help", "")]
    [string]$Action = ""
)

# --- Functions ---
function Write-Step        { Write-Host "`n>>> $($args -join ' ')" -ForegroundColor Cyan }
function Write-Success     { Write-Host "    [OK] $($args -join ' ')" -ForegroundColor Green }
function Write-Warning-C   { Write-Host "    [!!] $($args -join ' ')" -ForegroundColor Yellow }
function Write-Error-C     { Write-Host "    [!!] $($args -join ' ')" -ForegroundColor Red }
function Write-Info        { Write-Host "    [i] $($args -join ' ')" -ForegroundColor White }

$AddInName   = "kuroagent"
$Guid        = "14254940-5dfe-46ec-b860-a8291f526990"
$RegPath     = "HKCU:\Software\Microsoft\Office\Excel\Addins\$Guid"
$InstallDir  = Join-Path $env:USERPROFILE ".kuroagent"
$RepoUrl     = "https://github.com/yannassoumou/open-excel.git"

# --- Interactive menu ---
if ($Action -eq "") {
    Write-Host ""
    Write-Host "  ================================" -ForegroundColor Cyan
    Write-Host "  $AddInName Setup" -ForegroundColor Cyan
    Write-Host "  ================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1) Install   - First-time setup"
    Write-Host "  2) Update    - Pull latest and refresh"
    Write-Host "  3) Uninstall - Remove registry + CLI"
    Write-Host "  4) Purge     - Uninstall + delete .kuroagent folder"
    Write-Host ""
    $choice = Read-Host "  Choose [1-4]"
    switch ($choice) {
        "1" { $Action = "install" }
        "2" { $Action = "update" }
        "3" { $Action = "uninstall" }
        "4" { $Action = "purge" }
        default { Write-Error-C "Invalid choice. Aborting."; exit 1 }
    }
}

# ==================== INSTALL ====================
function Do-Install {
    Write-Step "Installing $AddInName"

    if (Test-Path (Join-Path $InstallDir ".git")) {
        Write-Step "Repository already exists at $InstallDir"
        Set-Location $InstallDir

        # Try a clean pull -- if it fails, nuke and re-clone
        $pullOk = $true
        $pullOut = & git fetch origin master 2>&1
        if ($LASTEXITCODE -eq 0) {
            $local  = & git rev-parse HEAD 2>$null
            $remote = & git rev-parse FETCH_HEAD 2>$null
            if ($local -ne $remote) {
                & git stash --include-untracked 2>&1 | Out-Null
                & git reset --hard FETCH_HEAD 2>&1 | Out-Null
                Write-Success "Pulled latest"
            } else {
                Write-Success "Already up to date"
            }
        } else {
            Write-Info "Pull failed, doing a clean reinstall"
            Set-Location $env:USERPROFILE
            & cmd /c "rmdir /s /q `"$InstallDir`"" 2>&1 | Out-Null
            if (Test-Path $InstallDir) {
                Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
            }
            $pullOk = $false
        }

        if (-not $pullOk -or -not (Test-Path (Join-Path $InstallDir "package.json"))) {
            Write-Step "Re-cloning repository ..."
            if (Test-Path $InstallDir) {
                Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
            }
            git clone $RepoUrl $InstallDir 2>&1 | Out-Null
            if (-not (Test-Path $InstallDir)) {
                Write-Error-C "Clone failed. Make sure git is installed and the repo is accessible."
                Write-Info "Install git: https://git-scm.com/download/win"
                exit 1
            }
            Write-Success "Cloned to $InstallDir"
        }

        Set-Location $InstallDir

        # Force re-link so we pick up the latest bin/kuroagent
        Write-Step "Refreshing global CLI link"
        & npm uninstall -g excel-custom-functions-js 2>&1 | Out-Null
    } else {
        # Clean up any leftover stub directory so clone works
        if (Test-Path $InstallDir) {
            Set-Location $env:USERPROFILE
            Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Step "Cloning repository ..."
        git clone $RepoUrl $InstallDir 2>&1 | Out-Null
        if (-not (Test-Path $InstallDir)) {
            Write-Error-C "Clone failed. Make sure git is installed and the repo is accessible."
            Write-Info "Install git: https://git-scm.com/download/win"
            exit 1
        }
        Write-Success "Cloned to $InstallDir"
        Set-Location $InstallDir
    }

    # Check Node.js
    Write-Step "Checking prerequisites"
    if (Get-Command node -ErrorAction SilentlyContinue) {
        Write-Success "Node.js $(& node -v)"
    } else {
        Write-Step "Node.js not found. Installing ..."
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            winget install OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        } else {
            $InstallerUrl = "https://nodejs.org/dist/v20.20.0/node-v20.20.0-x64.msi"
            $InstallerFile = "$env:TEMP\node-installer.msi"
            Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerFile -UseBasicParsing
            Start-Process msiexec.exe -ArgumentList "/i `"$InstallerFile`" /qn /norestart" -Wait
            Remove-Item $InstallerFile -Force
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        }
        Write-Success "Node.js installed: $(& node -v)"
    }

    Write-Success "npm $(& npm -v)"

    # Install dependencies
    Write-Step "Installing npm dependencies"
    $npmResult = & npm install --no-audit --no-fund --loglevel=error 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Error-C "npm install failed"
        $npmResult | ForEach-Object { Write-Info $_ }
        exit 1
    }
    Write-Success "Dependencies installed"

    # Dev certs
    Write-Step "Installing dev certificates"
    $certResult = & npx office-addin-dev-certs install --machine 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Dev certificates installed"
    } else {
        Write-Warning-C "Dev certs: already installed or non-zero exit"
    }

    # Register CLI
    Write-Step "Registering 'kuroagent' CLI"
    try {
        & npm link 2>$null | Out-Null
        Write-Success "npm link -- kuroagent CLI on PATH"
    } catch {
        $BinDir = Join-Path $env:APPDATA "npm"
        if (-not (Test-Path $BinDir)) { New-Item -ItemType Directory -Path $BinDir -Force | Out-Null }
        $CmdPath = Join-Path $BinDir "kuroagent.cmd"
        $NodePath = (Get-Command node).Source
        $ExcelScript = Join-Path $InstallDir "bin\kuroagent"
        Set-Content -Path $CmdPath -Value "@`"$NodePath`" `"$ExcelScript`" %"
        Write-Success "Created CLI wrapper at $CmdPath"
    }

    # Verify
    Write-Step "Verifying installation"
    if (Get-Command kuroagent -ErrorAction SilentlyContinue) {
        Write-Success "kuroagent CLI found"
        & kuroagent --version
    } else {
        Write-Error-C "kuroagent CLI not on PATH. Refresh terminal or restart."
        exit 1
    }

    # Manifest
    Write-Step "Checking manifest"
    try {
        & npx office-addin-manifest validate manifest.xml 2>$null | Out-Null
        Write-Success "manifest.xml valid"
    } catch {
        Write-Warning-C "Manifest validation returned non-zero (may be OK for dev)"
    }

    Show-Summary "install"
}

# ==================== UPDATE ====================
function Do-Update {
    Write-Step "Updating $AddInName"

    if (-not (Test-Path $InstallDir)) {
        Write-Error-C "Not installed. Run install first."
        exit 1
    }

    Set-Location $InstallDir

    # Re-link CLI after pull
    Write-Step "Pulling latest changes"
    try {
        git fetch origin master 2>&1 | Out-Null
        $local = git rev-parse HEAD 2>$null
        $remote = git rev-parse FETCH_HEAD 2>$null
        if ($local -eq $remote) {
            Write-Success "Already up to date"
        } else {
            git reset --hard FETCH_HEAD 2>&1 | Out-Null
            Write-Success "Pull complete"
        }
    } catch {
        git pull origin master
    }

    Write-Step "Reinstalling dependencies"
    & npm install --no-audit --no-fund --loglevel=error | Out-Null
    Write-Success "Dependencies updated"

    Write-Step "Re-linking kuroagent CLI"
    try {
        & npm uninstall -g kuroagent-custom-functions-js 2>$null | Out-Null
        & npm link 2>$null | Out-Null
        Write-Success "CLI re-linked"
    } catch {
        Write-Warning-C "npm link returned non-zero (may still work)"
    }

    # Verify
    if (Get-Command kuroagent -ErrorAction SilentlyContinue) {
        Write-Step "Verifying"
        Write-Success "kuroagent CLI: $(& kuroagent --version)"
    }

    Show-Summary "update"
}

# ==================== UNINSTALL ====================
function Do-Uninstall {
    param([bool]$RemoveDir = $false)

    Write-Step "Uninstalling $AddInName"

    # Registry
    Write-Step "Removing Excel registry sideloading"
    if (Test-Path $RegPath) {
        Remove-Item -Path $RegPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Success "Registry key removed"
    } else {
        Write-Warning-C "No registry key found (already clean)"
    }

    # CLI
    Write-Step "Unlinking kuroagent CLI"
    $NpmBin = Join-Path $env:APPDATA "npm"
    $RemovedAny = $false
    foreach ($f in @("kuroagent.cmd", "kuroagent", "kuroagent.ps1")) {
        $p = Join-Path $NpmBin $f
        if (Test-Path "$p") { Remove-Item "$p" -Force -ErrorAction SilentlyContinue; Write-Success "Removed: $f"; $RemovedAny = $true }
    }
    # Unlink via npm
    try { & npm uninstall -g excel-custom-functions-js 2>$null | Out-Null } catch {}
    if (-not $RemovedAny) { Write-Warning-C "CLI wrapper not found (may not have been linked)" }

    # Remove bin dir created by npm link
    $BinSubDir = Join-Path $NpmBin "kuroagent"
    if (Test-Path "$BinSubDir") {
        Remove-Item "$BinSubDir" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Success "Removed: kuroagent/ bin dir"
    }

    # Dev server
    Write-Step "Stopping dev server if running"
    try {
        $out = & cmd /c "netstat -ano | findstr :3000 | findstr LISTENING" 2>&1
        if ($out) {
            $pids = ($out -split "`n") | ForEach-Object {
                $parts = $_ -split "\s+"
                $parts[-1]
            } | Where-Object { $_ -match "^\d+$" }
            $pids | ForEach-Object { taskkill /F /PID $_ 2>$null }
            Write-Success "Stopped dev server on port 3000"
        } else {
            Write-Warning-C "No dev server running on port 3000"
        }
    } catch {
        Write-Warning-C "Could not check/stop server"
    }

    # Install dir
    if ($RemoveDir) {
        Write-Step "Removing installation directory: $InstallDir"
        if (Test-Path $InstallDir) {
            # cd out first so Windows lets us delete it
            Set-Location $env:USERPROFILE
            Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Success "Directory removed"
        }
    } else {
        Write-Warning-C "Keeping $InstallDir (use 'purge' to remove)"
    }

    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "  $AddInName uninstalled successfully" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  To reinstall:" -ForegroundColor Cyan
    Write-Host "    iwr https://raw.githubusercontent.com/yannassoumou/open-excel/master/setup.ps1 -OutFile setup.ps1"
    Write-Host "    powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    Write-Host ""
    Write-Host "  Note: Restart Excel to complete the uninstall." -ForegroundColor Yellow
    Write-Host ""
}

# ==================== SUMMARY ====================
function Show-Summary {
    param([string]$mode)
    if ($mode -eq "install") {
        $emoji = "installed"
    } else {
        $emoji = "updated"
    }
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "  $AddInName $emoji successfully" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Quick start:" -ForegroundColor Cyan
    Write-Host "    kuroagent                  Start dev server + sideload"
    Write-Host "    kuroagent -f file.xlsx     Start + open workbook"
    Write-Host "    kuroagent --stop           Stop dev server"
    Write-Host "    kuroagent -m path\xml      Use custom manifest"
    Write-Host "    kuroagent --no-open        Server only"
    Write-Host ""
    Write-Host "  Dev server runs on: https://localhost:3000" -ForegroundColor Cyan
    Write-Host "  Repo location:        $InstallDir" -ForegroundColor Cyan
    Write-Host ""
}

# ==================== HELP ====================
function Show-Help {
    Write-Host ""
    Write-Host "  $AddInName Unified Setup" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Usage:" -ForegroundColor White
    Write-Host "    setup.ps1              Interactive menu"
    Write-Host "    setup.ps1 install      First-time installation"
    Write-Host "    setup.ps1 update       Pull latest + refresh deps"
    Write-Host "    setup.ps1 uninstall    Remove registry + CLI"
    Write-Host "    setup.ps1 purge        Uninstall + delete .kuroagent"
    Write-Host ""
}

# ==================== DISPATCH ====================
switch ($Action) {
    "install"   { Do-Install }
    "update"    { Do-Update }
    "uninstall" { Do-Uninstall -RemoveDir $false }
    "purge"     { Do-Uninstall -RemoveDir $true }
    "help"      { Show-Help }
}
