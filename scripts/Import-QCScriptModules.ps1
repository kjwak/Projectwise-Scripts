# Dot-source from scripts/*.ps1: . (Join-Path $PSScriptRoot 'Import-QCScriptModules.ps1') -RepoRoot $repoRoot
param(
    [Parameter(Mandatory)]
    [string]$RepoRoot,
    [string[]]$AdditionalModules = @()
)

$ErrorActionPreference = 'Stop'
$modulesRoot = Join-Path $RepoRoot 'modules'
if (-not (Test-Path -LiteralPath $modulesRoot)) {
    throw "Modules folder not found: $modulesRoot. Use the Prepend PDF QC repo root (contains appsettings.json and modules\)."
}

$loadOrder = @(
    'Core.Results.psm1'
    'Core.Runtime.psm1'
    'Core.Database.psm1'
) + @($AdditionalModules)

foreach ($file in $loadOrder) {
    if ([string]::IsNullOrWhiteSpace($file)) { continue }
    $modPath = Join-Path $modulesRoot $file
    if (-not (Test-Path -LiteralPath $modPath)) {
        throw "Missing module file: $modPath"
    }
    Import-Module $modPath -Force -Scope Global -ErrorAction Stop | Out-Null
}

if (-not (Get-Command -Name 'Read-QCAppSettings' -ErrorAction SilentlyContinue)) {
    throw 'Core.Runtime did not load (Read-QCAppSettings unavailable). Verify modules\Core.Runtime.psm1 exists and is valid in this repo clone.'
}
