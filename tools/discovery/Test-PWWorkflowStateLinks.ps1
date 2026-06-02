<#
.SYNOPSIS
Live probe of Get-PWWorkflowStateLinks and Get-PWWorkflows for QC workflow validation.

.DESCRIPTION
Connects using projectWise settings from appsettings.json, calls Get-PWWorkflowStateLinks
with several invocation patterns, and reports state names / link properties so you can
compare against qcWorkflow.states (e.g. "Ready for QC").

Read-only. Does not change ProjectWise documents.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [string[]]$TargetStates = @('Ready for QC', 'In Production', 'Review In Progress'),
    [switch]$IncludeUnfilteredQuery,
    [int]$MaxSampleLinks = 5,
    [int]$QueryTimeoutSeconds = 90,
    [switch]$Pretty
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force -ErrorAction Stop
foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

function Get-LinkSnapshot {
    param([object]$Link)
    if (-not $Link) { return $null }
    $names = @(
        'WorkflowName', 'Workflow', 'StateName', 'State',
        'FromStateName', 'FromState', 'SourceStateName', 'SourceState', 'BeginState', 'BeginStateName',
        'ToStateName', 'ToState', 'TargetStateName', 'TargetState', 'EndState', 'EndStateName'
    )
    $h = [ordered]@{}
    foreach ($n in $names) {
        try {
            if ($Link.PSObject.Properties[$n] -and $null -ne $Link.$n) { $h[$n] = [string]$Link.$n }
        } catch { }
    }
    if ($h.Count -eq 0) {
        foreach ($p in $Link.PSObject.Properties) {
            if ($null -ne $p.Value -and $p.Value -isnot [System.Collections.IEnumerable]) {
                $h[$p.Name] = [string]$p.Value
            }
        }
    }
    return $h
}

function Get-UniqueStateLabels {
    param([object[]]$Links)
    $labels = [System.Collections.Generic.List[string]]::new()
    $propNames = @(
        'StateName', 'State', 'FromStateName', 'FromState', 'SourceStateName', 'SourceState',
        'ToStateName', 'ToState', 'TargetStateName', 'TargetState', 'BeginState', 'EndState', 'BeginStateName', 'EndStateName'
    )
    foreach ($link in @($Links)) {
        if (-not $link) { continue }
        foreach ($n in $propNames) {
            try {
                if ($link.PSObject.Properties[$n] -and -not [string]::IsNullOrWhiteSpace([string]$link.$n)) {
                    $v = ([string]$link.$n).Trim()
                    if (-not $labels.Contains($v)) { $labels.Add($v) | Out-Null }
                }
            } catch { }
        }
    }
    return @($labels | Sort-Object)
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
$pw = $config.projectWise
if (-not $pw) { throw 'appsettings.json has no projectWise section.' }

$ds = [string]$pw.datasourceName
$credPath = [string]$pw.credentialPath
if ([string]::IsNullOrWhiteSpace($ds)) { throw 'projectWise.datasourceName is empty.' }
if ([string]::IsNullOrWhiteSpace($credPath)) { throw 'projectWise.credentialPath is empty.' }

$qcw = $null
if ($config.ContainsKey('qcWorkflow') -and $config.qcWorkflow) {
    if ($config.qcWorkflow -is [hashtable]) { $qcw = $config.qcWorkflow }
    else {
        $qcw = @{}
        foreach ($p in $config.qcWorkflow.PSObject.Properties) { $qcw[$p.Name] = $p.Value }
    }
}
$workflowNames = @(
    if ($qcw) { [string]$qcw.expectedWorkflowName }; if ($qcw) { [string]$qcw.workflowName }
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

$configuredLabels = [System.Collections.Generic.List[string]]::new()
if ($qcw -and $qcw.states) {
    $st = $qcw.states
    if ($st -is [hashtable]) {
        foreach ($k in $st.Keys) { if ($st[$k]) { $configuredLabels.Add([string]$st[$k]) | Out-Null } }
    } elseif ($st.PSObject) {
        foreach ($p in $st.PSObject.Properties) { if ($p.Value) { $configuredLabels.Add([string]$p.Value) | Out-Null } }
    }
}
foreach ($k in @('stateAfterSuccessfulPrepend', 'defaultStateAfterPrepend', 'stateAfterFailedPrepend')) {
    if ($qcw -and $qcw.ContainsKey($k) -and $qcw[$k]) { $configuredLabels.Add([string]$qcw[$k]) | Out-Null }
}

$report = [ordered]@{
    appsettingsPath = $AppSettingsPath
    datasourceName  = $ds
    credentialPath    = $credPath
    configuredWorkflowNames = @($workflowNames)
    configuredStateLabels     = @($configuredLabels | Select-Object -Unique)
    targetStateChecks         = @()
    cmdletAvailable           = [bool](Get-Command -Name 'Get-PWWorkflowStateLinks' -ErrorAction SilentlyContinue)
    workflows                 = @()
    linkQueries               = @()
    errors                    = @()
}

$block = {
    $cmd = Get-Command -Name 'Get-PWWorkflowStateLinks' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return [ordered]@{ error = 'Get-PWWorkflowStateLinks not available after connect.' }
    }

    $wfCmd = Get-Command -Name 'Get-PWWorkflows' -ErrorAction SilentlyContinue
    $workflows = @()
    if ($wfCmd) {
        try { $workflows = @(Get-PWWorkflows -ErrorAction Stop) } catch {
            $script:wfErr = $_.Exception.Message
        }
    }

    $queries = [System.Collections.Generic.List[object]]::new()

    function Add-QueryResult {
        param([string]$Label, [hashtable]$QueryArgs, [object[]]$Links, [string]$ErrorMessage)
        $argCopy = @{}
        if ($QueryArgs) { foreach ($k in $QueryArgs.Keys) { $argCopy[$k] = $QueryArgs[$k] } }
        $stateLabels = @(Get-UniqueStateLabels -Links $Links)
        $linkFormat = if (@($Links).Count -eq 0) { 'none' }
            elseif ($stateLabels.Count -gt 0) { 'named' }
            elseif ($Links[0].PSObject.Properties['Id']) { 'id_chain' }
            else { 'unknown' }
        $queries.Add([ordered]@{
            label     = $Label
            arguments = $argCopy
            linkCount = @($Links).Count
            linkFormat = $linkFormat
            states    = @($stateLabels)
            samples   = @($Links | Select-Object -First $script:probeMaxSamples | ForEach-Object { Get-LinkSnapshot -Link $_ })
            error     = $ErrorMessage
        }) | Out-Null
    }

    # 1) Optional unfiltered query (can be very slow on large datasources)
    if ($script:probeIncludeUnfiltered) {
        try {
            $links = @(& $cmd -ErrorAction Stop)
        Add-QueryResult -Label 'no_parameters' -QueryArgs @{} -Links $links -ErrorMessage $null
    } catch {
        Add-QueryResult -Label 'no_parameters' -QueryArgs @{} -Links @() -ErrorMessage $_.Exception.Message
        }
    }

    # 2) Per configured workflow name
    foreach ($wn in $script:probeWorkflowNames) {
        foreach ($paramName in @('WorkflowName', 'Workflow')) {
            if (-not $cmd.Parameters.ContainsKey($paramName)) { continue }
            $invokeArgs = @{ $paramName = $wn }
            try {
                $links = @(& $cmd @invokeArgs -ErrorAction Stop)
                Add-QueryResult -Label ("${paramName}=$wn") -QueryArgs $invokeArgs -Links $links -ErrorMessage $null
            } catch {
                Add-QueryResult -Label ("${paramName}=$wn") -QueryArgs $invokeArgs -Links @() -ErrorMessage $_.Exception.Message
            }
        }
    }

    return [ordered]@{
        workflows = @($workflows | ForEach-Object {
            $n = $null
            foreach ($p in @('Name', 'WorkflowName', 'Workflow')) {
                try { if ($_.PSObject.Properties[$p]) { $n = [string]$_.$p; break } } catch { }
            }
            [ordered]@{ name = $n; properties = @(Get-LinkSnapshot -Link $_) }
        })
        linkQueries = @($queries)
        workflowListError = $script:wfErr
    }
}

try {
    $script:probeWorkflowNames = @($workflowNames)
    $script:probeIncludeUnfiltered = $IncludeUnfilteredQuery.IsPresent
    $script:probeMaxSamples = [Math]::Max(1, $MaxSampleLinks)
    $live = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -ScriptBlock $block
    if ($live.error) {
        $report.errors += [string]$live.error
    } else {
        $report.workflows = @($live.workflows)
        $report.linkQueries = @($live.linkQueries)
        if ($live.workflowListError) { $report.errors += ('Get-PWWorkflows: ' + [string]$live.workflowListError) }
    }
} catch {
    $report.errors += $_.Exception.Message
}

foreach ($target in @($TargetStates)) {
    $foundIn = @()
    foreach ($q in @($report.linkQueries)) {
        foreach ($s in @($q.states)) {
            if ([string]$s -eq [string]$target) {
                $foundIn += [string]$q.label
            }
        }
    }
    $inConfig = [bool](@($report.configuredStateLabels | Where-Object { $_.Equals($target, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0)
    $report.targetStateChecks += [ordered]@{
        targetState       = [string]$target
        inQcWorkflowStates = [bool]$inConfig
        foundInLinkQueries = @($foundIn | Select-Object -Unique)
        exactMatchAnyQuery = ([bool](@($foundIn).Count -gt 0))
    }
}

$json = $report | ConvertTo-Json -Depth 12
if ($Pretty.IsPresent) {
    $json | Write-Output
} else {
  $json
}
