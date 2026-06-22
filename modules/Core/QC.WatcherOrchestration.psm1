Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$orchRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $orchRoot 'Core.Results.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $orchRoot 'Core.Runtime.psm1') -Force -WarningAction SilentlyContinue

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

function Get-QCInitiatedWorkflowStateName {
    <#
    .SYNOPSIS
    Resolved display name for the QC Initiated workflow state from config (default: QC Initiated).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            if (-not (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue)) {
                Import-Module (Join-Path $orchRoot 'QC.Workflow.psm1') -Force -WarningAction SilentlyContinue | Out-Null
            }
            $wf = Get-QCWorkflowSettings -Config $Config
            $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcInitiated'
            if (-not [string]::IsNullOrWhiteSpace($resolved)) { return [string]$resolved }
        } catch { }
    }
    try {
        if ($Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
            $wf = $Config.qcWorkflow
            if ($wf -is [hashtable] -and $wf.ContainsKey('states') -and $wf.states) {
                $st = $wf.states
                if ($st -is [hashtable] -and $st.ContainsKey('qcInitiated') -and $st.qcInitiated) {
                    return [string]$st.qcInitiated
                }
                if ($st.qcInitiated) { return [string]$st.qcInitiated }
            }
        }
    } catch { }
    return 'QC Initiated'
}

function Test-QCWorkflowStateIsQcInitiated {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StateName,
        [Parameter(Mandatory)][hashtable]$Config
    )
    if ([string]::IsNullOrWhiteSpace($StateName)) { return $false }
    $initiated = Get-QCInitiatedWorkflowStateName -Config $Config
    return ($StateName.Trim().ToLowerInvariant() -eq $initiated.Trim().ToLowerInvariant())
}

function Get-QCFinalizingWorkflowStateName {
    <#
    .SYNOPSIS
    Resolved display name for the QC Finalizing workflow state from config (default: QC Finalizing).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            if (-not (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue)) {
                Import-Module (Join-Path $orchRoot 'QC.Workflow.psm1') -Force -WarningAction SilentlyContinue | Out-Null
            }
            $wf = Get-QCWorkflowSettings -Config $Config
            $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcFinalizing'
            if (-not [string]::IsNullOrWhiteSpace($resolved)) { return [string]$resolved }
        } catch { }
    }
    try {
        if ($Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
            $wf = $Config.qcWorkflow
            if ($wf -is [hashtable] -and $wf.ContainsKey('states') -and $wf.states) {
                $st = $wf.states
                if ($st -is [hashtable] -and $st.ContainsKey('qcFinalizing') -and $st.qcFinalizing) {
                    return [string]$st.qcFinalizing
                }
                if ($st.qcFinalizing) { return [string]$st.qcFinalizing }
            }
        }
    } catch { }
    return 'QC Finalizing'
}

function Test-QCWorkflowStateIsQcFinalizing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StateName,
        [Parameter(Mandatory)][hashtable]$Config
    )
    if ([string]::IsNullOrWhiteSpace($StateName)) { return $false }
    $finalizing = Get-QCFinalizingWorkflowStateName -Config $Config
    return ($StateName.Trim().ToLowerInvariant() -eq $finalizing.Trim().ToLowerInvariant())
}

function Get-QCReadyForVerificationWorkflowStateName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            if (-not (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue)) {
                Import-Module (Join-Path $orchRoot 'QC.Workflow.psm1') -Force -WarningAction SilentlyContinue | Out-Null
            }
            $wf = Get-QCWorkflowSettings -Config $Config
            $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'readyForVerification'
            if (-not [string]::IsNullOrWhiteSpace($resolved)) { return [string]$resolved }
        } catch { }
    }
    try {
        if ($Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
            $wf = $Config.qcWorkflow
            if ($wf -is [hashtable] -and $wf.ContainsKey('states') -and $wf.states) {
                $st = $wf.states
                if ($st -is [hashtable] -and $st.ContainsKey('readyForVerification') -and $st.readyForVerification) {
                    return [string]$st.readyForVerification
                }
                if ($st.readyForVerification) { return [string]$st.readyForVerification }
            }
        }
    } catch { }
    return 'Ready for Verification'
}

function Test-QCWorkflowStateIsReadyForVerification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StateName,
        [Parameter(Mandatory)][hashtable]$Config
    )
    if ([string]::IsNullOrWhiteSpace($StateName)) { return $false }
    $ready = Get-QCReadyForVerificationWorkflowStateName -Config $Config
    return ($StateName.Trim().ToLowerInvariant() -eq $ready.Trim().ToLowerInvariant())
}

function Test-QCWorkflowStateIsAutomationIntake {
    <#
    .SYNOPSIS
    Automation-owned intake states that enqueue prepend and suppress user notifications.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StateName,
        [Parameter(Mandatory)][hashtable]$Config
    )
    if (Test-QCWorkflowStateIsQcInitiated -StateName $StateName -Config $Config) { return $true }
    if (Test-QCWorkflowStateIsQcFinalizing -StateName $StateName -Config $Config) { return $true }
    return $false
}

function Get-QCPrependAuditActions {
    <#
    .SYNOPSIS
    Audit action names that may enqueue QC_PREPEND for paired sheet PDFs (QC Initiated state and/or QC_Archivist tag).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $defaults = @(
        'DOCUMENT_MODIFY'
        'DOCUMENT_ATTR'
        'DOCUMENT_CIN'
        'DOCUMENT_FILE_REP'
        'DOCUMENT_VERSION'
        'DOCUMENT_CREATE'
        'DOCUMENT_STATE'
    )
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = $Config.auditPoller
            if ($ap -is [hashtable] -and $ap.ContainsKey('qcPrependAuditActions') -and $ap.qcPrependAuditActions) {
                $list = @($ap.qcPrependAuditActions | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($list.Count -gt 0) { return $list }
            }
            elseif ($ap.PSObject -and $ap.qcPrependAuditActions) {
                $list = @($ap.qcPrependAuditActions | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($list.Count -gt 0) { return $list }
            }
        }
    } catch { }
    return $defaults
}

function Invoke-QCRecoverQueue {
    <#
    .SYNOPSIS
    RecoverQueue: stale/orphan running jobs, duplicate repair, optional watcher-active clear.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [switch]$ClearWatcherActive
    )
    if (Get-Command -Name 'Invoke-QCQueueStartupCheck' -ErrorAction SilentlyContinue) {
        return Invoke-QCQueueStartupCheck -Config $Config -ClearWatcherActive:$ClearWatcherActive.IsPresent
    }
    return New-QCFailureResult -Code 'RECOVER_QUEUE_UNAVAILABLE' -Message 'QC.Queue.Json not loaded.' -Data @{}
}

function Invoke-QCReconcileAudit {
    <#
    .SYNOPSIS
    ReconcileAudit: ingest audit events from watermark minus restart overlap; does not run folder scans.
    Requires an open PW session and Invoke-AuditTrailScan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$WatermarkPath,
        [array]$WatchRootConfigs = @(),
        [switch]$DryRun
    )

    $stats = @{ eventsFetched = 0; dbWrites = 0; candidates = 0; watermarkBefore = $null; watermarkAfter = $null; durationMs = 0 }
    if (-not (Get-Command -Name 'Get-AuditTrailPollWindow' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'RECONCILE_AUDIT_UNAVAILABLE' -Message 'PW.AuditPoller not loaded.' -Data $stats
    }
    $pollWindow = Get-AuditTrailPollWindow -Config $Config -WatermarkPath $WatermarkPath -UseRestartOverlap
    $stats.watermarkBefore = $pollWindow.watermarkBefore
    if ($DryRun.IsPresent) {
        return New-QCSuccessResult -Code 'RECONCILE_AUDIT_DRYRUN' -Message 'Audit reconcile dry-run (no PW query).' -Data @{
            sinceUtc = $pollWindow.sinceUtc
            untilUtc = $pollWindow.untilUtc
            watermarkBefore = $pollWindow.watermarkBefore
            restartOverlapUsed = [bool]$pollWindow.restartOverlapUsed
        }
    }
    $scan = Invoke-AuditTrailScan -Config $Config -Since $pollWindow.since -Until $pollWindow.until -WatchRootConfigs $WatchRootConfigs
    if (-not $scan.IsSuccess) {
        return New-QCFailureResult -Code 'RECONCILE_AUDIT_SCAN_FAILED' -Message $scan.Message -Data $stats
    }
    $d = $scan.Data
    if ($d.stats) {
        $stats.eventsFetched = [int]$d.stats.totalEvents
        $stats.dbWrites = [int]$d.stats.dbWrites
        $stats.durationMs = [int]$d.durationMs
    }
    if ($d.candidates) { $stats.candidates = @($d.candidates).Count }
    $stats.watermarkAfter = [string]$d.watermarkAfter
    return New-QCSuccessResult -Code 'RECONCILE_AUDIT_OK' -Message 'Startup audit reconcile completed.' -Data @{
        stats = $stats
        pollWindow = $pollWindow
        scan = $d
    }
}

function Invoke-QCReconcileOutputs {
    <#
    .SYNOPSIS
    ReconcileOutputs: lightweight verification of recent succeeded jobs and pending notification work.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $out = @{
        notificationJobsPending = 0
        recentSucceededJobs = 0
        warnings = @()
    }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'RECONCILE_OUTPUTS_SKIPPED' -Message 'Database disabled.' -Data $out
    }
    try {
        $res = Invoke-QCDatabaseScalar -Config $Config -Sql @"
SELECT COUNT(*) FROM processing_jobs
WHERE job_type = 'QC_NOTIFICATION' AND status IN ('pending','running')
"@
        if ($res.IsSuccess) { $out.notificationJobsPending = [int]$res.Data.value }
    } catch {
        $out.warnings += $_.Exception.Message
    }
    try {
        $root = $null
        if ($Config.queue -and $Config.queue.rootDir) { $root = [string]$Config.queue.rootDir }
        if ($root -and (Test-Path -LiteralPath (Join-Path $root 'succeeded'))) {
            $out.recentSucceededJobs = @(Get-ChildItem -LiteralPath (Join-Path $root 'succeeded') -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
        }
    } catch { }
    return New-QCSuccessResult -Code 'RECONCILE_OUTPUTS_OK' -Message 'Output reconcile snapshot complete.' -Data $out
}

function Get-QCAuditWatermarkAgeSeconds {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [string]$WatermarkPath = '')
    $wm = $null
    if (Get-Command -Name 'Get-AuditTrailCaptureWatermark' -ErrorAction SilentlyContinue) {
        $wm = Get-AuditTrailCaptureWatermark -Config $Config -WatermarkPath $WatermarkPath
    }
    if (-not $wm) { return $null }
    return [int][Math]::Max(0, ((Get-Date).ToUniversalTime() - $wm.ToUniversalTime()).TotalSeconds)
}

function Invoke-QCWatcherStartupSequence {
    <#
    .SYNOPSIS
    Startup: RecoverQueue → ReconcileAudit → ReconcileOutputs. Returns telemetry for WATCH_STARTUP_SEQUENCE log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$QueueRoot,
        [array]$WatchRootConfigs = @(),
        [switch]$ClearWatcherActive,
        [switch]$DryRun,
        [switch]$SkipAuditReconcile
    )

    $watermarkPath = Join-Path (Join-Path $QueueRoot '_watcher') 'audit-capture-watermark.txt'
    $telemetry = @{
        queueRecovery = $null
        auditReconcile = $null
        outputReconcile = $null
        watermarkAgeSeconds = $null
        errors = [System.Collections.Generic.List[string]]::new()
    }

    $rec = Invoke-QCRecoverQueue -Config $Config -ClearWatcherActive:$ClearWatcherActive.IsPresent
    if ($rec.IsSuccess) { $telemetry.queueRecovery = $rec.Data } else { [void]$telemetry.errors.Add([string]$rec.Message) }

    try { $telemetry.watermarkAgeSeconds = Get-QCAuditWatermarkAgeSeconds -Config $Config -WatermarkPath $watermarkPath } catch { }

    if (-not $SkipAuditReconcile.IsPresent) {
        $auditRec = Invoke-QCReconcileAudit -Config $Config -WatermarkPath $watermarkPath -WatchRootConfigs $WatchRootConfigs -DryRun:$DryRun.IsPresent
        if ($auditRec.IsSuccess) {
            $telemetry.auditReconcile = $auditRec.Data
            if (-not $DryRun.IsPresent -and $auditRec.Data.scan) {
                $scanData = $auditRec.Data.scan
                $fetched = 0
                try {
                    if ($scanData.stats -and $null -ne $scanData.stats.totalEvents) { $fetched = [int]$scanData.stats.totalEvents }
                } catch { }
                if ($fetched -gt 0 -and $scanData.watermarkAfter) {
                    $captured = $null
                    try {
                        $captured = [DateTime]::ParseExact(
                            ([string]$scanData.watermarkAfter).Trim().TrimEnd('Z'),
                            'yyyy-MM-dd HH:mm:ss',
                            $null,
                            [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
                        )
                    } catch { }
                    if ($captured) {
                        [void](Set-AuditTrailCaptureWatermark -WatermarkPath $watermarkPath -CapturedThrough $captured -Config $Config)
                    }
                }
            }
        } else {
            [void]$telemetry.errors.Add([string]$auditRec.Message)
        }
    }

    $outRec = Invoke-QCReconcileOutputs -Config $Config
    if ($outRec.IsSuccess) { $telemetry.outputReconcile = $outRec.Data } else { [void]$telemetry.errors.Add([string]$outRec.Message) }

    $code = if ($telemetry.errors.Count -gt 0) { 'WATCHER_STARTUP_PARTIAL' } else { 'WATCHER_STARTUP_OK' }
    return New-QCSuccessResult -Code $code -Message 'Watcher startup sequence completed.' -Data $telemetry
}

function _QCWO-GetConfigSectionHashtable {
    param([hashtable]$Config, [string]$Key)
    if (-not $Config -or -not $Config.ContainsKey($Key) -or -not $Config[$Key]) { return $null }
    $sec = $Config[$Key]
    if ($sec -is [hashtable]) { return $sec }
    if (Get-Command -Name 'ConvertTo-HashtableDeep' -ErrorAction SilentlyContinue) {
        return ConvertTo-HashtableDeep -Value $sec
    }
    return $null
}

function Get-QCFullScanScheduleTimesFromConfig {
    <#
    .SYNOPSIS
    Wall-clock times (HH:mm) for scheduled full PW folder scans. Uses auditPoller.fullScanSchedule.times, then reconciliation.folderScanSchedule.times.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $raw = @()
    foreach ($loc in @(
            @{ section = 'auditPoller'; scheduleKey = 'fullScanSchedule' },
            @{ section = 'reconciliation'; scheduleKey = 'folderScanSchedule' }
        )) {
        $sec = _QCWO-GetConfigSectionHashtable -Config $Config -Key $loc.section
        if (-not $sec) { continue }
        $sched = $null
        if ($sec.ContainsKey($loc.scheduleKey) -and $sec[$loc.scheduleKey]) {
            $sched = $sec[$loc.scheduleKey]
        }
        if (-not $sched) { continue }
        $list = $null
        if ($sched -is [hashtable] -and $sched.ContainsKey('times')) { $list = $sched.times }
        elseif ($sched.PSObject -and $sched.times) { $list = $sched.times }
        if ($list) { $raw = @($list) }
        if ($raw.Count -gt 0) { break }
    }

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $raw) {
        $text = ([string]$t).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        if ($text -notmatch '^(\d{1,2}):(\d{2})$') { continue }
        $h = [int]$Matches[1]
        $m = [int]$Matches[2]
        if ($h -lt 0 -or $h -gt 23 -or $m -lt 0 -or $m -gt 59) { continue }
        $hhmm = '{0:D2}:{1:D2}' -f $h, $m
        if ($normalized -notcontains $hhmm) { [void]$normalized.Add($hhmm) }
    }
    return @($normalized | Sort-Object)
}

function _QCWO-GetFullScanScheduleStatePath {
    param([hashtable]$Config, [string]$QueueRoot)
    $root = $QueueRoot
    if ([string]::IsNullOrWhiteSpace($root)) {
        $q = _QCWO-GetConfigSectionHashtable -Config $Config -Key 'queue'
        if ($q -and $q.ContainsKey('rootDir') -and $q.rootDir) { $root = [string]$q.rootDir }
    }
    if ([string]::IsNullOrWhiteSpace($root)) { return $null }
    return Join-Path (Join-Path $root '_watcher') 'full-scan-schedule-state.json'
}

function Test-QCFullScanScheduleSlotComplete {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$SlotKey,
        [string]$QueueRoot = ''
    )
    if (Get-Command -Name 'Get-QCWatcherStateValue' -ErrorAction SilentlyContinue) {
        $v = Get-QCWatcherStateValue -Config $Config -StateKey $SlotKey
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $true }
    }
    $path = _QCWO-GetFullScanScheduleStatePath -Config $Config -QueueRoot $QueueRoot
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return $false }
    try {
        $doc = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($doc.completedSlots) {
            return @($doc.completedSlots | ForEach-Object { [string]$_ }) -contains $SlotKey
        }
    } catch { }
    return $false
}

function Set-QCFullScanScheduleSlotComplete {
    <#
    .SYNOPSIS
    Marks a full-scan schedule slot (full_scan_schedule|yyyy-MM-dd|HH:mm) complete for today.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$SlotKey,
        [string]$QueueRoot = ''
    )
    $completedAt = (Get-Date).ToUniversalTime().ToString('o')
    $ok = $false
    if (Get-Command -Name 'Set-QCWatcherStateValue' -ErrorAction SilentlyContinue) {
        $ok = [bool](Set-QCWatcherStateValue -Config $Config -StateKey $SlotKey -StateValue $completedAt)
    }
    $path = _QCWO-GetFullScanScheduleStatePath -Config $Config -QueueRoot $QueueRoot
    if ($path) {
        try {
            $dir = Split-Path -Parent $path
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $slots = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            if (Test-Path -LiteralPath $path) {
                $doc = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($s in @($doc.completedSlots)) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$s)) { [void]$slots.Add([string]$s) }
                }
            }
            [void]$slots.Add($SlotKey)
            $payload = @{ completedSlots = @($slots | Sort-Object) }
            Set-Content -LiteralPath $path -Value ($payload | ConvertTo-Json -Compress) -Encoding UTF8
            $ok = $true
        } catch { }
    }
    return $ok
}

function Get-QCFullFolderScanReconciliationPlan {
    <#
    .SYNOPSIS
    Returns whether a scheduled full folder scan is due (wall-clock times in display time zone).
    Falls back to reconcileEveryNCycles when no schedule times are configured.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Nullable[int]]$CycleNum = $null,
        [string]$QueueRoot = ''
    )

    $none = @{
        due = $false
        mode = 'none'
        reason = $null
        scheduledTime = $null
        slotKey = $null
        scheduledTimes = @()
        reconcileEvery = $null
    }

    $times = Get-QCFullScanScheduleTimesFromConfig -Config $Config
    if ($times.Count -gt 0) {
        Set-QCDisplayTimeZoneFromConfig -Config $Config
        $now = Get-QCWallClockNow
        $day = $now.Date.ToString('yyyy-MM-dd')
        foreach ($hhmm in $times) {
            $parts = $hhmm.Split(':')
            $slotAt = $now.Date.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
            if ($now -lt $slotAt) { continue }
            $slotKey = "full_scan_schedule|$day|$hhmm"
            if (-not (Test-QCFullScanScheduleSlotComplete -Config $Config -SlotKey $slotKey -QueueRoot $QueueRoot)) {
                return @{
                    due = $true
                    mode = 'schedule'
                    reason = 'scheduled_time'
                    scheduledTime = $hhmm
                    slotKey = $slotKey
                    scheduledTimes = $times
                    reconcileEvery = $null
                }
            }
        }
        return @{
            due = $false
            mode = 'schedule'
            reason = 'all_slots_complete_today'
            scheduledTime = $null
            slotKey = $null
            scheduledTimes = $times
            reconcileEvery = $null
        }
    }

    $reconcileEvery = 0
    $ap = _QCWO-GetConfigSectionHashtable -Config $Config -Key 'auditPoller'
    if ($ap -and $ap.ContainsKey('reconcileEveryNCycles') -and $ap.reconcileEveryNCycles) {
        try { $reconcileEvery = [int]$ap.reconcileEveryNCycles } catch { $reconcileEvery = 0 }
    }
    if ($reconcileEvery -lt 1 -or -not $CycleNum) { return $none }
    $dueCycle = ($CycleNum -ge $reconcileEvery) -and (($CycleNum % $reconcileEvery) -eq 0)
    if (-not $dueCycle) { return $none }
    return @{
        due = $true
        mode = 'cycle'
        reason = 'scheduled_cycle'
        scheduledTime = $null
        slotKey = $null
        scheduledTimes = @()
        reconcileEvery = $reconcileEvery
    }
}

function _QCWO-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[[string]$k] = $Value[$k] }
        return $h
    }
    if ($Value -is [pscustomobject]) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function Get-QCWatcherStallRecoverySettings {
    <#
    .SYNOPSIS
    Dashboard-side recovery when the Watch-QCTrigger child stops writing logs (blocked PW SQL, etc.).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $defaults = @{
        enabled = $true
        noLogActivitySeconds = 600
        auditScanMaxSeconds = 300
        sendSessionAlert = $true
        postKillCooldownSeconds = 120
    }

    $settings = @{}
    foreach ($key in @($defaults.Keys)) { $settings[$key] = $defaults[$key] }

    $watcher = _QCWO-ToHashtable $Config.watcher
    $stall = $null
    if ($watcher -and $watcher.ContainsKey('stallRecovery')) {
        $stall = _QCWO-ToHashtable $watcher.stallRecovery
    }
    if ($stall) {
        if ($stall.ContainsKey('enabled') -and $null -ne $stall.enabled) { try { $settings.enabled = [bool]$stall.enabled } catch { } }
        if ($stall.ContainsKey('noLogActivitySeconds') -and $null -ne $stall.noLogActivitySeconds) { try { $settings.noLogActivitySeconds = [int]$stall.noLogActivitySeconds } catch { } }
        if ($stall.ContainsKey('auditScanMaxSeconds') -and $null -ne $stall.auditScanMaxSeconds) { try { $settings.auditScanMaxSeconds = [int]$stall.auditScanMaxSeconds } catch { } }
        if ($stall.ContainsKey('sendSessionAlert') -and $null -ne $stall.sendSessionAlert) { try { $settings.sendSessionAlert = [bool]$stall.sendSessionAlert } catch { } }
        if ($stall.ContainsKey('postKillCooldownSeconds') -and $null -ne $stall.postKillCooldownSeconds) { try { $settings.postKillCooldownSeconds = [int]$stall.postKillCooldownSeconds } catch { } }
    }

    if ($settings.noLogActivitySeconds -lt 120) { $settings.noLogActivitySeconds = 120 }
    if ($settings.auditScanMaxSeconds -lt 60) { $settings.auditScanMaxSeconds = 60 }
    if ($settings.postKillCooldownSeconds -lt 30) { $settings.postKillCooldownSeconds = 30 }

    return $settings
}

function Test-QCWatcherChildStalled {
    <#
    .SYNOPSIS
    Returns whether a live watcher child appears wedged (no JSONL progress / audit scan never finished).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings,
        [Parameter(Mandatory)]
        [bool]$WatcherAlive,
        [Nullable[datetime]]$LastLogActivityUtc,
        [string]$LastWatcherEventCode = '',
        [Nullable[datetime]]$LastWatcherEventUtc,
        [Nullable[datetime]]$WatcherSpawnedAtUtc,
        [string]$CurrentScanStage = '',
        [Nullable[datetime]]$StageSinceUtc,
        [Nullable[datetime]]$NowUtc = $null
    )

    if (-not [bool]$Settings.enabled -or -not $WatcherAlive) {
        return @{ stalled = $false; reason = 'not_applicable' }
    }

    $now = if ($NowUtc) { $NowUtc.ToUniversalTime() } else { (Get-Date).ToUniversalTime() }
    $lastLog = if ($LastLogActivityUtc) { $LastLogActivityUtc.ToUniversalTime() } else { $null }
    $lastEvent = if ($LastWatcherEventUtc) { $LastWatcherEventUtc.ToUniversalTime() } else { $null }
    $spawnedAt = if ($WatcherSpawnedAtUtc) { $WatcherSpawnedAtUtc.ToUniversalTime() } else { $null }
    $stageSince = if ($StageSinceUtc) { $StageSinceUtc.ToUniversalTime() } else { $null }

    $code = ([string]$LastWatcherEventCode).Trim().ToUpperInvariant()
    $stage = ([string]$CurrentScanStage).Trim()

    if ($spawnedAt -and $lastEvent -and $lastEvent -lt $spawnedAt) {
        $lastEvent = $null
        $code = ''
    }
    if ($spawnedAt -and $stageSince -and $stageSince -lt $spawnedAt) {
        $stageSince = $null
        $stage = ''
    }

    if ($spawnedAt -and $lastLog -and $lastLog -lt $spawnedAt) {
        $lastLog = $spawnedAt
    }

    if ($code -eq 'WATCH_AUDIT_SCAN_START' -and $lastEvent) {
        $auditSec = ($now - $lastEvent).TotalSeconds
        if ($auditSec -ge [double]$Settings.auditScanMaxSeconds) {
            return @{
                stalled = $true
                reason = 'audit_scan_timeout'
                seconds = [int][Math]::Floor($auditSec)
                thresholdSeconds = [int]$Settings.auditScanMaxSeconds
                lastEventCode = $code
            }
        }
    }

    if ($stage -match '(?i)audit trail scan starting' -and $stageSince) {
        $stageSec = ($now - $stageSince).TotalSeconds
        if ($stageSec -ge [double]$Settings.auditScanMaxSeconds) {
            return @{
                stalled = $true
                reason = 'audit_scan_stage_timeout'
                seconds = [int][Math]::Floor($stageSec)
                thresholdSeconds = [int]$Settings.auditScanMaxSeconds
                lastEventCode = $code
            }
        }
    }

    if ($lastLog) {
        $silentSec = ($now - $lastLog).TotalSeconds
        if ($silentSec -ge [double]$Settings.noLogActivitySeconds) {
            return @{
                stalled = $true
                reason = 'no_log_activity'
                seconds = [int][Math]::Floor($silentSec)
                thresholdSeconds = [int]$Settings.noLogActivitySeconds
                lastEventCode = $code
            }
        }
    }

    return @{ stalled = $false; reason = 'healthy' }
}

function Stop-QCWatcherChildForStall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Process]$Process,
        [string]$Reason = 'watcher_child_stalled',
        [int]$SecondsSilent = 0
    )

    $pidVal = 0
    try { $pidVal = [int]$Process.Id } catch { }
    $killed = $false
    $exitCode = $null
    $alreadyExited = $false
    try {
        if ($Process.HasExited) {
            $alreadyExited = $true
            try { $exitCode = [int]$Process.ExitCode } catch { }
        } else {
            $Process.Kill()
            $killed = $true
            try { $Process.WaitForExit(15000) } catch { }
            try { $exitCode = [int]$Process.ExitCode } catch { }
        }
    } catch {
        return @{
            killed = $false
            pid = $pidVal
            reason = $Reason
            secondsSilent = $SecondsSilent
            alreadyExited = $alreadyExited
            errorMessage = [string]$_.Exception.Message
        }
    }

    return @{
        killed = $killed
        pid = $pidVal
        reason = $Reason
        secondsSilent = $SecondsSilent
        exitCode = $exitCode
        alreadyExited = $alreadyExited
        errorMessage = $null
    }
}

Export-ModuleMember -Function Get-QCWatcherMode, Get-QCReconciliationPlan, Get-QCReconcileStatusSetsOnStart, Get-QCWatcherContinuousSettings, Get-QCInitiatedWorkflowStateName, Test-QCWorkflowStateIsQcInitiated, Get-QCFinalizingWorkflowStateName, Test-QCWorkflowStateIsQcFinalizing, Get-QCReadyForVerificationWorkflowStateName, Test-QCWorkflowStateIsReadyForVerification, Test-QCWorkflowStateIsAutomationIntake, Get-QCPrependAuditActions, Invoke-QCRecoverQueue, Invoke-QCReconcileAudit, Invoke-QCReconcileOutputs, Invoke-QCWatcherStartupSequence, Get-QCAuditWatermarkAgeSeconds, Get-QCFullScanScheduleTimesFromConfig, Get-QCFullFolderScanReconciliationPlan, Test-QCFullScanScheduleSlotComplete, Set-QCFullScanScheduleSlotComplete, Get-QCWatcherStallRecoverySettings, Test-QCWatcherChildStalled, Stop-QCWatcherChildForStall
