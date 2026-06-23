function Resolve-ModuleImplPath {
    <#
    .SYNOPSIS
    Resolves a bare module file name to its Phase 4 folder implementation path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [string]$ModulesDir
    )

    if (-not $ModulesDir) {
        $ModulesDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules'
    }

    $rel = ($ModuleName -replace '/', '\').Trim()
    if ($rel -match '[\\/]') {
        $impl = Join-Path $ModulesDir $rel
        if (-not (Test-Path -LiteralPath $impl)) {
            throw "Module implementation not found: $impl"
        }
        return $impl
    }

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

    if ($map.ContainsKey($rel)) {
        $impl = Join-Path $ModulesDir $map[$rel]
    } else {
        $impl = Join-Path $ModulesDir $rel
    }

    if (-not (Test-Path -LiteralPath $impl)) {
        throw "Module implementation not found: $impl"
    }
    return $impl
}

function Get-QCModuleImplementation {
    <#
    .SYNOPSIS
    Returns the loaded implementation module when present in the session.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [string]$ModulesDir
    )

    $implPath = [System.IO.Path]::GetFullPath((Resolve-ModuleImplPath -ModuleName $ModuleName -ModulesDir $ModulesDir))
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ModuleName)
    foreach ($m in @(Get-Module -Name $baseName -All)) {
        if ($m.Path -and ([System.IO.Path]::GetFullPath($m.Path) -eq $implPath)) {
            return $m
        }
    }

    return (Get-Module -Name $baseName | Select-Object -First 1)
}

function Remove-QCModuleFlatShims {
    <#
    .SYNOPSIS
    No-op after Phase 4H shim removal; kept for test compatibility.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [string]$ModulesDir
    )
}
