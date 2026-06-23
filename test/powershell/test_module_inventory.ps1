# Ensures modules/FILES.md lists every *.psm1 (and no extras).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesDir = Join-Path $repoRoot 'modules'
$filesMd = Join-Path $modulesDir 'FILES.md'

$onDisk = @(Get-ChildItem -LiteralPath $modulesDir -Filter '*.psm1' | ForEach-Object { $_.Name } | Sort-Object)
$raw = Get-Content -LiteralPath $filesMd -Raw
$inDoc = [regex]::Matches($raw, '`([A-Za-z0-9._-]+\.psm1)`') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique

$missingFromDoc = @($onDisk | Where-Object { $_ -notin $inDoc })
$extraInDoc = @($inDoc | Where-Object { $_ -notin $onDisk })

if ($missingFromDoc.Count -gt 0 -or $extraInDoc.Count -gt 0) {
    if ($missingFromDoc.Count -gt 0) {
        Write-Host "Missing from modules/FILES.md: $($missingFromDoc -join ', ')"
    }
    if ($extraInDoc.Count -gt 0) {
        Write-Host "Listed in FILES.md but not on disk: $($extraInDoc -join ', ')"
    }
    throw 'modules/FILES.md is out of sync with modules/*.psm1. Update FILES.md and modules/README.md when adding modules.'
}

Write-Host "test_module_inventory.ps1 passed ($($onDisk.Count) modules)"
