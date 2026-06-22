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

function _QCResolve-ModuleRelativePath {
    param([Parameter(Mandatory)][string]$FileName)
    if ($FileName -match '[\\/]') { return $FileName }
    $map = @{
        'Core.Results.psm1' = 'Core\Core.Results.psm1'
        'Core.Runtime.psm1' = 'Core\Core.Runtime.psm1'
        'Core.Paths.psm1' = 'Core\Core.Paths.psm1'
        'Core.Config.psm1' = 'Core\Core.Config.psm1'
        'Core.Logging.psm1' = 'Core\Core.Logging.psm1'
        'Core.Hashing.psm1' = 'Core\Core.Hashing.psm1'
        'Core.Telemetry.psm1' = 'Core\Core.Telemetry.psm1'
        'QC.WatcherOrchestration.psm1' = 'Core\QC.WatcherOrchestration.psm1'
        'Core.Database.psm1' = 'Database\Core.Database.psm1'
        'PW.Connection.psm1' = 'ProjectWise\PW.Connection.psm1'
        'PW.Discovery.psm1' = 'ProjectWise\PW.Discovery.psm1'
        'PW.AuditPoller.psm1' = 'ProjectWise\PW.AuditPoller.psm1'
        'PW.Users.psm1' = 'ProjectWise\PW.Users.psm1'
        'QC.Workflow.psm1' = 'Workflow\QC.Workflow.psm1'
        'QC.AuditTriggers.psm1' = 'Workflow\QC.AuditTriggers.psm1'
        'QC.ProcessType.psm1' = 'Workflow\QC.ProcessType.psm1'
        'QC.Queue.Json.psm1' = 'Queue\QC.Queue.Json.psm1'
        'QC.JobFactory.psm1' = 'Queue\QC.JobFactory.psm1'
        'QC.Worker.psm1' = 'Queue\QC.Worker.psm1'
        'QC.Filters.psm1' = 'Queue\QC.Filters.psm1'
        'QC.Triggers.psm1' = 'Queue\QC.Triggers.psm1'
        'QC.Processors.psm1' = 'Processing\QC.Processors.psm1'
        'QC.StatusSet.psm1' = 'Processing\QC.StatusSet.psm1'
        'QC.Rendition.psm1' = 'Processing\QC.Rendition.psm1'
        'QC.ReviewStamp.psm1' = 'Processing\QC.ReviewStamp.psm1'
        'QC.PdfExport.psm1' = 'Processing\QC.PdfExport.psm1'
        'QC.Notifications.psm1' = 'Notifications\QC.Notifications.psm1'
        'QC.NotificationGraph.psm1' = 'Notifications\QC.NotificationGraph.psm1'
        'QC.Reporting.psm1' = 'Reporting\QC.Reporting.psm1'
        'QC.DebugMcp.psm1' = 'Diagnostics\QC.DebugMcp.psm1'
    }
    if ($map.ContainsKey($FileName)) { return $map[$FileName] }
    return $FileName
}

function _QCImport-ModuleFile {
    param([Parameter(Mandatory)][string]$Path)
    Import-Module $Path -Force -Scope Global -WarningAction SilentlyContinue | Out-Null
}

# Core.Database pulls Runtime/Results into module-only scope and can hide globals on Windows PS 5.1.
$loadOrder = @(
    'Core\Core.Results.psm1'
    'Database\Core.Database.psm1'
    'Core\Core.Runtime.psm1'
    'Core\Core.Results.psm1'
) + @($AdditionalModules)

$seen = @{}
foreach ($file in $loadOrder) {
    if ([string]::IsNullOrWhiteSpace($file)) { continue }
    if ($seen.ContainsKey($file)) { continue }
    $seen[$file] = $true
    $rel = _QCResolve-ModuleRelativePath -FileName $file
    $modPath = Join-Path $modulesRoot $rel
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
