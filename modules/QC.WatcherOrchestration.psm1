Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-QCWatcherMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [switch]$ReconcileStatusSetsFirst
    )
    $mode = 'audit_only'
    try {
        if ($Config.ContainsKey('watcher') -and $Config.watcher -and $Config.watcher.mode) {
            $mode = ([string]$Config.watcher.mode).Trim().ToLowerInvariant()
        }
    } catch { }
    if ($ReconcileStatusSetsFirst.IsPresent -and $mode -eq 'audit_only') { $mode = 'hybrid' }
    if ($mode -notin @('audit_only','reconciliation','recovery','hybrid')) { $mode = 'audit_only' }
    return $mode
}

function Get-QCReconciliationPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$WatcherMode,
        [Nullable[datetime]]$LastSuccessfulAuditWatermark,
        [Nullable[datetime]]$NowUtc = (Get-Date).ToUniversalTime()
    )
    $enabled = $true
    $triggerSource = 'none'
    $reason = $null
    $run = $false
    $downtimeSeconds = 0
    $auditGapDetected = $false
    $lastReconUtc = $null
    $maxDowntime = 0
    try {
        if ($Config.ContainsKey('reconciliation') -and $Config.reconciliation) {
            $rc = $Config.reconciliation
            if ($rc.ContainsKey('enabled')) { $enabled = [bool]$rc.enabled }
            if ($rc.ContainsKey('lastRunUtc') -and $rc.lastRunUtc) { $lastReconUtc = [datetime]::Parse([string]$rc.lastRunUtc).ToUniversalTime() }
            if ($rc.ContainsKey('downtimeThresholdSeconds') -and $rc.downtimeThresholdSeconds) { $maxDowntime = [int]$rc.downtimeThresholdSeconds }
        }
    } catch { }

    if (-not $enabled) {
        return @{ shouldRun = $false; reason = 'disabled'; triggerSource = 'config'; downtimeSeconds = 0; auditGapDetected = $false; lastReconciliationUtc = $lastReconUtc }
    }
    if ($WatcherMode -eq 'reconciliation') { return @{ shouldRun = $true; reason = 'explicit_mode'; triggerSource = 'manual'; downtimeSeconds = 0; auditGapDetected = $false; lastReconciliationUtc = $lastReconUtc } }
    if ($WatcherMode -in @('hybrid','recovery')) {
        if ($LastSuccessfulAuditWatermark) {
            $downtimeSeconds = [int][Math]::Max(0, ($NowUtc.Value - $LastSuccessfulAuditWatermark.Value).TotalSeconds)
        }
        if ($maxDowntime -gt 0 -and $downtimeSeconds -ge $maxDowntime) {
            $run = $true; $reason = 'downtime_threshold'; $triggerSource = 'scheduler'; $auditGapDetected = $true
        } else {
            $run = $false
            $reason = if ($WatcherMode -eq 'hybrid') { 'hybrid_audit_first' } else { 'recovery_no_gap' }
            $triggerSource = 'none'
            $auditGapDetected = $false
        }
        return @{ shouldRun = $run; reason = $reason; triggerSource = $triggerSource; downtimeSeconds = $downtimeSeconds; auditGapDetected = $auditGapDetected; lastReconciliationUtc = $lastReconUtc }
    }
    return @{ shouldRun = $false; reason = 'audit_only'; triggerSource = 'none'; downtimeSeconds = 0; auditGapDetected = $false; lastReconciliationUtc = $lastReconUtc }
}

Export-ModuleMember -Function Get-QCWatcherMode, Get-QCReconciliationPlan
