# Ensures modules/FILES.md documents every organized module implementation.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesDir = Join-Path $repoRoot 'modules'
$filesMd = Join-Path $modulesDir 'FILES.md'

$onDisk = @(Get-ChildItem -LiteralPath $modulesDir -Filter '*.psm1' -File -Recurse |
    Where-Object { $_.DirectoryName -ne $modulesDir } |
    ForEach-Object { $_.FullName.Substring($modulesDir.Length + 1).Replace('\', '/') } |
    Sort-Object)
$raw = Get-Content -LiteralPath $filesMd -Raw
$inDoc = @([regex]::Matches($raw, '(?m)^\| `([\w./-]+\.psm1)` \|') |
    ForEach-Object { $_.Groups[1].Value } |
    Sort-Object)

$differences = @(Compare-Object -ReferenceObject $onDisk -DifferenceObject $inDoc)
foreach ($difference in $differences) {
    if ($difference.SideIndicator -eq '<=') {
        Write-Host "Missing from modules/FILES.md: $($difference.InputObject)"
    } else {
        Write-Host "Unexpected or duplicate path in modules/FILES.md: $($difference.InputObject)"
    }
}
if ($differences.Count -gt 0) {
    throw 'modules/FILES.md is out of sync with the organized module implementations.'
}

Write-Host "test_module_inventory.ps1 passed ($($onDisk.Count) modules)"
