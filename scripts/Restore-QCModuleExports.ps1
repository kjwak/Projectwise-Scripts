# Dot-source from script entrypoints:
#   . (Join-Path $PSScriptRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
#
# Restores Windows PowerShell 5.1 script-session exports after nested Import-Module -Force
# in feature modules drops foundation commands (Core.Runtime, Core.Database, etc.).

param(
    [Parameter(Mandatory)]
    [string]$RepoRoot
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

if (-not $script:QCModuleBootstrapRepoRoot) {
    $script:QCModuleBootstrapRepoRoot = $RepoRoot
}
$script:QCModuleBootstrapModulesRoot = Join-Path $RepoRoot 'modules'

$script:QCFoundationRestoreOrder = @(
    'Core\Core.Results.psm1'
    'Core\Core.Paths.psm1'
    'Core\Core.Runtime.psm1'
    'Core\Core.Config.psm1'
    'Core\Core.Logging.psm1'
    'Core\Core.Hashing.psm1'
    'Database\Core.Database.psm1'
)

$script:QCModulePathMap = @{
    'Core.Results.psm1' = 'Core\Core.Results.psm1'
    'Core.Runtime.psm1' = 'Core\Core.Runtime.psm1'
    'Core.Paths.psm1' = 'Core\Core.Paths.psm1'
    'Core.Config.psm1' = 'Core\Core.Config.psm1'
    'Core.Logging.psm1' = 'Core\Core.Logging.psm1'
    'Core.Hashing.psm1' = 'Core\Core.Hashing.psm1'
    'Core.Telemetry.psm1' = 'Core\Core.Telemetry.psm1'
    'QC.WatcherOrchestration.psm1' = 'Core\QC.WatcherOrchestration.psm1'
    'QC.StatusSetBatching.psm1' = 'Core\QC.StatusSetBatching.psm1'
    'Core.Database.psm1' = 'Database\Core.Database.psm1'
    'QC.Queue.Json.psm1' = 'Queue\QC.Queue.Json.psm1'
    'QC.HostThrottle.psm1' = 'Queue\QC.HostThrottle.psm1'
    'QC.JobFactory.psm1' = 'Queue\QC.JobFactory.psm1'
    'QC.Worker.psm1' = 'Queue\QC.Worker.psm1'
    'QC.Filters.psm1' = 'Queue\QC.Filters.psm1'
    'QC.Triggers.psm1' = 'Queue\QC.Triggers.psm1'
    'QC.Processors.psm1' = 'Processing\QC.Processors.psm1'
    'QC.StatusSet.psm1' = 'Processing\QC.StatusSet.psm1'
    'QC.Rendition.psm1' = 'Processing\QC.Rendition.psm1'
    'QC.Notifications.psm1' = 'Notifications\QC.Notifications.psm1'
    'QC.WatcherAlerts.psm1' = 'Notifications\QC.WatcherAlerts.psm1'
    'QC.NotificationGraph.psm1' = 'Notifications\QC.NotificationGraph.psm1'
    'PW.Connection.psm1' = 'ProjectWise\PW.Connection.psm1'
    'PW.Discovery.psm1' = 'ProjectWise\PW.Discovery.psm1'
    'PW.AuditPoller.psm1' = 'ProjectWise\PW.AuditPoller.psm1'
    'PW.Users.psm1' = 'ProjectWise\PW.Users.psm1'
    'QC.Workflow.psm1' = 'Workflow\QC.Workflow.psm1'
    'QC.ProcessType.psm1' = 'Workflow\QC.ProcessType.psm1'
    'QC.DebugMcp.psm1' = 'Diagnostics\QC.DebugMcp.psm1'
}

function Resolve-QCModulePath {
    <#
    .SYNOPSIS
    Resolves a folder implementation module path under modules/ (never flat shims).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath,
        [string]$RepoRoot = $script:QCModuleBootstrapRepoRoot
    )

    $rel = ($RelativePath -replace '/', '\').Trim()
    if ($rel -match '[\\/]') {
        return Join-Path (Join-Path $RepoRoot 'modules') $rel
    }
    if ($script:QCModulePathMap.ContainsKey($rel)) {
        return Join-Path (Join-Path $RepoRoot 'modules') $script:QCModulePathMap[$rel]
    }
    return Join-Path (Join-Path $RepoRoot 'modules') $rel
}

function Import-QCModuleGlobal {
    <#
    .SYNOPSIS
    Imports a folder implementation module with -Force -Global (session export restore).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath,
        [string]$RepoRoot = $script:QCModuleBootstrapRepoRoot
    )

    $path = Resolve-QCModulePath -RelativePath $RelativePath -RepoRoot $RepoRoot
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing module file: $path"
    }
    Import-Module $path -Force -Global -WarningAction SilentlyContinue | Out-Null
}

function Restore-QCFoundationModuleExports {
    <#
    .SYNOPSIS
    Re-imports foundation modules in standard order to restore script-session exports.
    #>
    [CmdletBinding()]
    param(
        [string]$RepoRoot = $script:QCModuleBootstrapRepoRoot
    )

    foreach ($rel in $script:QCFoundationRestoreOrder) {
        Import-QCModuleGlobal -RelativePath $rel -RepoRoot $RepoRoot
    }
}

function Test-QCRequiredCommands {
    <#
    .SYNOPSIS
    Throws if any named commands are not available in the current session.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Names,
        [string]$Context = 'module bootstrap'
    )

    $missing = @($Names | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
        })
    if ($missing.Count -gt 0) {
        throw ("{0} failed: missing command(s): {1}" -f $Context, ($missing -join ', '))
    }
}

function Import-QCModuleBootstrapSet {
    <#
    .SYNOPSIS
    Imports feature modules, restores foundation exports, then validates required commands.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$FeatureModules,
        [string[]]$RequiredCommands = @(),
        [string]$Context = 'module bootstrap',
        [string]$RepoRoot = $script:QCModuleBootstrapRepoRoot,
        [switch]$SkipFoundationRestore
    )

    foreach ($rel in @($FeatureModules)) {
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }
        Import-QCModuleGlobal -RelativePath $rel -RepoRoot $RepoRoot
    }
    if (-not $SkipFoundationRestore.IsPresent) {
        Restore-QCFoundationModuleExports -RepoRoot $RepoRoot
    }
    if ($RequiredCommands -and $RequiredCommands.Count -gt 0) {
        Test-QCRequiredCommands -Names $RequiredCommands -Context $Context
    }
}
