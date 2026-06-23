# Sync tools/overlay/qc_review_stamp.py into the PyInstaller onedir bundle without a full rebuild.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$src = Join-Path $repoRoot 'tools\overlay\qc_review_stamp.py'
$dst = Join-Path $repoRoot 'dist\qc_overlay_prepend\_internal\qc_review_stamp.py'
if (-not (Test-Path -LiteralPath $src)) { throw "Source not found: $src" }
$dstDir = Split-Path -Parent $dst
if (-not (Test-Path -LiteralPath $dstDir)) { throw "Bundle folder not found: $dstDir (run tools\overlay\build_overlay_exe.ps1 first)" }
Copy-Item -LiteralPath $src -Destination $dst -Force
Write-Host "Synced: $dst"
