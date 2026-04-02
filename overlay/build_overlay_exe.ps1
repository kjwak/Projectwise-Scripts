# build_overlay_exe.ps1
#
# Builds dist\qc_overlay_prepend\qc_overlay_prepend.exe (one-folder) for machines that do NOT have Python
# (e.g. ProjectWise automation hosts). Run this script only on a build machine where Python is installed.
#
# Your ProjectWise / trigger PowerShell should invoke the exe directly, for example:
#   & "C:\Path\to\qc_overlay_prepend\qc_overlay_prepend.exe" $incomingPdf $qcHistoryPdf -o $outputPdf
#
# Usage (from repo): .\overlay\build_overlay_exe.ps1
# Output: dist\qc_overlay_prepend\qc_overlay_prepend.exe

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$venvPath = Join-Path $projectRoot ".venv_overlay_build"
$specFile = Join-Path $projectRoot "qc_overlay_prepend.spec"

# Resolve Python (Windows often has 'py' launcher but not 'python' on PATH)
$pythonCmd = $null
if (Get-Command python -ErrorAction SilentlyContinue) { $pythonCmd = "python" }
elseif (Get-Command py -ErrorAction SilentlyContinue) { $pythonCmd = "py" }
if (-not $pythonCmd) { throw "Python not found. Install Python and ensure 'python' or 'py' is on PATH." }

if (-not (Test-Path $specFile)) {
    throw "Spec not found: $specFile"
}

Push-Location $projectRoot

try {
    if (-not (Test-Path $venvPath)) {
        Write-Host "Creating build venv..."
        & $pythonCmd -m venv $venvPath
    }
    $pythonExe = Join-Path $venvPath "Scripts\python.exe"
    $pipExe = Join-Path $venvPath "Scripts\pip.exe"

    Write-Host "Installing PyInstaller and overlay dependencies (overlay/requirements.txt)..."
    & $pipExe install -r (Join-Path $scriptDir "requirements.txt") pyinstaller --quiet

    Write-Host "Building qc_overlay_prepend (onedir) from qc_overlay_prepend.spec (may take 2-3 minutes)..."
    & $pythonExe -m PyInstaller --clean --noconfirm $specFile

    $exe = Join-Path $projectRoot "dist\qc_overlay_prepend\qc_overlay_prepend.exe"
    if (Test-Path $exe) {
        $size = (Get-Item $exe).Length / 1MB
        Write-Host ""
        Write-Host "SUCCESS: $exe"
        Write-Host "Size: $([math]::Round($size, 2)) MB"
        Write-Host ""
        Write-Host "Deploy: copy the entire dist\qc_overlay_prepend folder to the target PC (exe needs sibling _internal)."
        Write-Host "prepend_qc.ps1 defaults to dist\qc_overlay_prepend.exe (onefile); use -QcOverlayExe for dist\qc_overlay_prepend\qc_overlay_prepend.exe if you deploy the onedir folder."
        Write-Host "Automation: call the exe from PowerShell (e.g. after a ProjectWise trigger):"
        Write-Host '  & "C:\path\to\qc_overlay_prepend\qc_overlay_prepend.exe" incoming.pdf qc_history.pdf -o output.pdf'
    } else {
        Write-Error "Build failed - exe not found at $exe"
    }
} finally {
    Pop-Location
}
