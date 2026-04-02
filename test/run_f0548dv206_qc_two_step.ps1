# run_f0548dv206_qc_two_step.ps1
# Two-step QC overlay prepend test for sheet f0548dv206:
#   Step 1: incoming = Version 01, history = Version 00
#   Step 2: incoming = Version 02, history = output of step 1
#
# Usage (from repo root):
#   .\test\run_f0548dv206_qc_two_step.ps1
# Optional:
#   .\test\run_f0548dv206_qc_two_step.ps1 -ExePath "C:\path\to\qc_overlay_prepend.exe"

param(
    [string]$ExePath = ""
)

$ErrorActionPreference = "Stop"
$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $testDir

. (Join-Path $projectRoot "Resolve-OverlayExe.ps1")

if (-not $ExePath) {
    $ExePath = Select-ExistingOverlayExePath $projectRoot
    if (-not $ExePath) {
        $ExePath = Join-Path $projectRoot "dist\qc_overlay_prepend\qc_overlay_prepend.exe"
    }
}

if (-not (Test-Path -LiteralPath $ExePath)) {
    throw "Executable not found: $ExePath`nBuild with: .\overlay\build_overlay_exe.ps1 (outputs dist\qc_overlay_prepend\ with _internal)"
}

$ExePath = Resolve-OverlayExePath $ExePath

$v00 = Join-Path $testDir "Version 00\f0548dv206.pdf"
$v01 = Join-Path $testDir "Version 01\f0548dv206.pdf"
$v02 = Join-Path $testDir "Version 02\f0548dv206.pdf"
$outDir = Join-Path $testDir "output"
$step1 = Join-Path $outDir "f0548dv206_qc_v00_to_v01.pdf"
$step2 = Join-Path $outDir "f0548dv206_qc_then_v02.pdf"
$currentMaster = Join-Path $outDir "f0548dv206_current_master.pdf"

foreach ($p in @($v00, $v01, $v02)) {
    if (-not (Test-Path $p)) {
        throw "Missing input PDF: $p"
    }
}

if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

# Remove previous outputs so the exe only creates new files (no in-process delete/replace).
foreach ($p in @($step1, $step2)) {
    if (Test-Path $p) {
        Remove-Item -Force $p
    }
}

# Reset current-master to Version 00 for deterministic step-1 baseline:
# Step 1 must compare Old=V00 (red) vs New=V01 (green+current).
if (Test-Path $currentMaster) {
    Remove-Item -Force $currentMaster
}
Copy-Item -Path $v00 -Destination $currentMaster -Force

Write-Host "=== Step 1: incoming=Version 01, history=Version 00 ===" -ForegroundColor Cyan
Write-Host "  -> $step1"
& $ExePath $v01 $v00 -o $step1 --current-master $currentMaster
$e1 = $LASTEXITCODE
Write-Host "Exit code: $e1"
if ($e1 -ne 0) { exit $e1 }

Write-Host ""
Write-Host "=== Step 2: incoming=Version 02, history=step 1 output ===" -ForegroundColor Cyan
Write-Host "  -> $step2"
& $ExePath $v02 $step1 -o $step2 --current-master $currentMaster
$e2 = $LASTEXITCODE
Write-Host "Exit code: $e2"
if ($e2 -ne 0) { exit $e2 }

Write-Host ""
Write-Host "=== Page counts ===" -ForegroundColor Green
$py = Join-Path $projectRoot ".venv_overlay_build\Scripts\python.exe"
if (Test-Path $py) {
    & $py -c @"
import pikepdf
from pathlib import Path
base = Path(r'$outDir')
for name in ['f0548dv206_qc_v00_to_v01.pdf', 'f0548dv206_qc_then_v02.pdf']:
    p = base / name
    with pikepdf.open(p) as d:
        print(f'  {name}: {len(d.pages)} page(s)')
"@
} else {
    Write-Host "  (Install venv at .venv_overlay_build or open PDFs manually to verify page count.)"
}

Write-Host ""
Write-Host "Done. Outputs:" -ForegroundColor Green
Write-Host "  $step1"
Write-Host "  $step2"
Write-Host "  $currentMaster"
