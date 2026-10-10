#Requires -Version 5.1
<#
.SYNOPSIS
    Offline-friendly Python venv setup for a portable local Python tree.

.DESCRIPTION
    Assumes this layout (relative to the script, or current directory):

        .
        ├── python\          # Portable Python (contains python.exe)
        │   └── python.exe
        ├── app.py
        ├── requirements.txt
        ├── wheels\          # Created if missing; holds downloaded wheels
        ├── .venv\           # Created by this script
        ├── setup.ps1        # This script
        └── run.bat          # Created by this script

    Steps:
      1. Create .venv with the local python.exe
      2. Activate the venv
      3. If wheels\ is empty / missing, download packages into wheels\
      4. pip install --no-index --find-links=wheels from requirements.txt
      5. Write run.bat that activates the venv and runs app.py
#>

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Paths (everything relative to the folder that contains this script)
# ---------------------------------------------------------------------------
$Root        = $PSScriptRoot
if (-not $Root) { $Root = (Get-Location).Path }

$PythonDir   = Join-Path $Root 'python'
$PythonExe   = Join-Path $PythonDir 'python.exe'
$VenvDir     = Join-Path $Root '.venv'
$VenvPython  = Join-Path $VenvDir 'Scripts\python.exe'
$VenvPip     = Join-Path $VenvDir 'Scripts\pip.exe'
$WheelsDir   = Join-Path $Root 'wheels'
$ReqFile     = Join-Path $Root 'requirements.txt'
$AppFile     = Join-Path $Root 'app.py'
$RunBat      = Join-Path $Root 'run.bat'
$ActivatePs1 = Join-Path $VenvDir 'Scripts\Activate.ps1'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Step([string]$Message) {
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Assert-File([string]$Path, [string]$Hint) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file not found: $Path`n$Hint"
    }
}

# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------
Write-Step 'Checking local Python and project files'
Assert-File $PythonExe "Place a portable Python distribution in: $PythonDir"
Assert-File $ReqFile   "Create requirements.txt next to this script."
Assert-File $AppFile   "Create app.py next to this script."

Write-Host "  python.exe : $PythonExe"
Write-Host "  requirements.txt : $ReqFile"
Write-Host "  app.py : $AppFile"

# ---------------------------------------------------------------------------
# 1. Create virtual environment
# ---------------------------------------------------------------------------
Write-Step 'Creating virtual environment (.venv)'
if (Test-Path -LiteralPath $VenvDir) {
    Write-Host "  .venv already exists – reusing it."
} else {
    & $PythonExe -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) { throw "Failed to create venv (exit $LASTEXITCODE)" }
    Write-Host "  Created: $VenvDir"
}

Assert-File $VenvPython "venv was created but Scripts\python.exe is missing."

# ---------------------------------------------------------------------------
# 2. Activate venv (for the rest of this script)
# ---------------------------------------------------------------------------
Write-Step 'Activating virtual environment'
. $ActivatePs1
Write-Host "  Active Python: $((Get-Command python).Source)"

# ---------------------------------------------------------------------------
# 3. Ensure wheels directory exists; download if empty
# ---------------------------------------------------------------------------
Write-Step 'Preparing local wheels directory'
if (-not (Test-Path -LiteralPath $WheelsDir)) {
    New-Item -ItemType Directory -Path $WheelsDir | Out-Null
    Write-Host "  Created: $WheelsDir"
}

$existingWheels = @(Get-ChildItem -LiteralPath $WheelsDir -Filter *.whl -ErrorAction SilentlyContinue)
$needDownload = ($existingWheels.Count -eq 0)

if ($needDownload) {
    Write-Host "  No wheels found – downloading packages into wheels\"
    # Use the venv pip so we get the correct platform tags
    & $VenvPip download `
        -r $ReqFile `
        -d $WheelsDir `
        --prefer-binary
    if ($LASTEXITCODE -ne 0) {
        throw "pip download failed (exit $LASTEXITCODE). Check network / requirements.txt."
    }
    $downloaded = @(Get-ChildItem -LiteralPath $WheelsDir -Filter *.whl)
    Write-Host "  Downloaded $($downloaded.Count) wheel(s)."
} else {
    Write-Host "  Found $($existingWheels.Count) existing wheel(s) – skipping download."
}

# ---------------------------------------------------------------------------
# 4. Offline install from local wheels
# ---------------------------------------------------------------------------
Write-Step 'Installing packages offline from wheels\'
& $VenvPip install `
    --no-index `
    --find-links $WheelsDir `
    -r $ReqFile
if ($LASTEXITCODE -ne 0) {
    throw "Offline pip install failed (exit $LASTEXITCODE). Ensure all required wheels are present."
}
Write-Host "  Packages installed successfully."

# ---------------------------------------------------------------------------
# 5. Create run.bat
# ---------------------------------------------------------------------------
Write-Step 'Creating run.bat'
$batContent = @"
@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo Virtual environment not found. Run setup.ps1 first.
    exit /b 1
)

call ".venv\Scripts\activate.bat"
python app.py %*
"@

Set-Content -LiteralPath $RunBat -Value $batContent -Encoding ASCII
Write-Host "  Written: $RunBat"

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
Write-Host "`n========================================" -ForegroundColor Green
Write-Host " Setup complete." -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host @"

Usage:
  1. Double-click run.bat   (or run it from a cmd/PowerShell window)
  2. Or manually:
       .\.venv\Scripts\Activate.ps1
       python app.py

To refresh wheels later (e.g. after changing requirements.txt):
  Remove-Item .\wheels\* -Force
  .\setup.ps1
"@
