Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-QCReconcileStatusSetsOnStart {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $onStart = $true
    try {
        if ($Config.ContainsKey('reconciliation') -and $Config.reconciliation) {
            $rc = $Config.reconciliation
            if ($rc.ContainsKey('reconcileStatusSetsOnStart')) { return [bool]$rc.reconcileStatusSetsOnStart }
        }
        if ($Config.ContainsKey('statusSet') -and $Config.statusSet -and $Config.statusSet.ContainsKey('reconcileOnWatcherStart')) {
            return [bool]$Config.statusSet.reconcileOnWatcherStart
        }
    } catch { }
    return $onStart
}

function Get-QCWatcherContinuousSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [switch]$ContinuousSwitch,
        [Nullable[int]]$PollIntervalMsOverride
    )
    $continuous = $false
    $pollSleepMs = 750
    try {
        if ($Config.ContainsKey('workers') -and $Config.workers -and $Config.workers.idleSleepMs) {
            $pollSleepMs = [int]$Config.workers.idleSleepMs
        }
        if ($Config.ContainsKey('watcher') -and $Config.watcher) {
            $w = $Config.watcher
            if ($w.ContainsKey('continuous')) { $continuous = [bool]$w.continuous }
            if ($w.ContainsKey('idleSleepMs') -and $null -ne $w.idleSleepMs) { $pollSleepMs = [int]$w.idleSleepMs }
        }
    } catch { }
    if ($ContinuousSwitch.IsPresent) { $continuous = $true }
    if ($null -ne $PollIntervalMsOverride -and $PollIntervalMsOverride -gt 0) { $pollSleepMs = [int]$PollIntervalMsOverride }
    if ($pollSleepMs -lt 100) { $pollSleepMs = 100 }
    return @{ continuous = $continuous; pollSleepMs = $pollSleepMs }
}

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
    $reconcileOnStart = Get-QCReconcileStatusSetsOnStart -Config $Config
    if ($ReconcileStatusSetsFirst.IsPresent -and $reconcileOnStart -and $mode -eq 'audit_only') { $mode = 'hybrid' }
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

Export-ModuleMember -Function Get-QCWatcherMode, Get-QCReconciliationPlan, Get-QCReconcileStatusSetsOnStart, Get-QCWatcherContinuousSettings
