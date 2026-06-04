# Dot-source from scripts/*.ps1: . (Join-Path $PSScriptRoot 'Import-QCScriptModules.ps1') -RepoRoot $repoRoot
param(
    [Parameter(Mandatory)]
    [string]$RepoRoot,
    [string[]]$AdditionalModules = @()
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$modulesRoot = Join-Path $RepoRoot 'modules'
if (-not (Test-Path -LiteralPath $modulesRoot)) {
    throw "Modules folder not found: $modulesRoot. Use the Prepend PDF QC repo root (contains appsettings.json and modules\)."
}

function _QCImport-ModuleFile {
    param([Parameter(Mandatory)][string]$Path)
    Import-Module $Path -Force -Scope Global -WarningAction SilentlyContinue | Out-Null
}

# Core.Database pulls Runtime/Results into module-only scope and can hide globals on Windows PS 5.1.
$loadOrder = @(
    'Core.Results.psm1'
    'Core.Database.psm1'
    'Core.Runtime.psm1'
    'Core.Results.psm1'
) + @($AdditionalModules)

$seen = @{}
foreach ($file in $loadOrder) {
    if ([string]::IsNullOrWhiteSpace($file)) { continue }
    if ($seen.ContainsKey($file)) { continue }
    $seen[$file] = $true
    $modPath = Join-Path $modulesRoot $file
    if (-not (Test-Path -LiteralPath $modPath)) {
        throw "Missing module file: $modPath"
    }
    _QCImport-ModuleFile -Path $modPath
}

if (-not (Get-Command -Name 'Read-QCAppSettings' -ErrorAction SilentlyContinue)) {
    function Read-QCAppSettings {
        [CmdletBinding()]
        param([Parameter(Mandatory)][string]$Path)
        & (Get-Command -Name 'Read-QCAppSettings' -Module 'Core.Runtime') -Path $Path
    }
}

if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
    function Test-QCDatabaseEnabled {
        [CmdletBinding()]
        param([Parameter(Mandatory)][hashtable]$Config)
        & (Get-Command -Name 'Test-QCDatabaseEnabled' -Module 'Core.Database') -Config $Config
    }
}

if (-not (Get-Command -Name 'Initialize-QCDatabaseSchema' -ErrorAction SilentlyContinue)) {
    function Initialize-QCDatabaseSchema {
        [CmdletBinding()]
        param([Parameter(Mandatory)][hashtable]$Config)
        & (Get-Command -Name 'Initialize-QCDatabaseSchema' -Module 'Core.Database') -Config $Config
    }
}

if (-not (Get-Command -Name 'Invoke-QCDatabaseScalar' -ErrorAction SilentlyContinue)) {
    function Invoke-QCDatabaseScalar {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][hashtable]$Config,
            [Parameter(Mandatory)][string]$Sql,
            [hashtable]$Parameters = @{},
            [int]$CommandTimeout = -1
        )
        & (Get-Command -Name 'Invoke-QCDatabaseScalar' -Module 'Core.Database') @PSBoundParameters
    }
}

if (-not (Get-Command -Name 'Invoke-QCDatabaseNonQuery' -ErrorAction SilentlyContinue)) {
    function Invoke-QCDatabaseNonQuery {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][hashtable]$Config,
            [Parameter(Mandatory)][string]$Sql,
            [hashtable]$Parameters = @{},
            [int]$CommandTimeout = -1
        )
        & (Get-Command -Name 'Invoke-QCDatabaseNonQuery' -Module 'Core.Database') @PSBoundParameters
    }
}

if (-not (Get-Command -Name 'Read-QCAppSettings' -ErrorAction SilentlyContinue)) {
    $loaded = (Get-Module | ForEach-Object { $_.Name }) -join ', '
    throw "Read-QCAppSettings is still unavailable after import. Loaded modules: $loaded"
}
