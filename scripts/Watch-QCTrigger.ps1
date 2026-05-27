<#
.SYNOPSIS
One-shot filesystem watcher tick: detect → classify → enqueue.

.DESCRIPTION
Loads appsettings.json, scans configured watchFolders once, applies:
  - Core.Paths (normalize)
  - QC.Filters (allow/deny)
  - QC.Triggers (match/no match)
  - QC.JobFactory (build job + dedupe key)
  - QC.Queue.Json (enqueue)

Constraints:
  - No ProjectWise writes
  - No processor execution
  - Run-once only (no continuous loop)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = '',

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [int]$MaxFiles = 0,

    # When set, the watcher first walks every workspace under statusSet.localRoot
    # and reconciles its local _StatusSet.pdf back to ProjectWise (legacy parity:
    # every restart re-checks every manifest). Pass this on the first invocation
    # after restart; omit it from subsequent watcher ticks to skip the re-walk.
    [Parameter(Mandatory = $false)]
    [switch]$ReconcileStatusSetsFirst
)

$ErrorActionPreference = 'Stop'
$watchRunSw = [System.Diagnostics.Stopwatch]::StartNew()
$phaseMs = @{}
$phaseCounts = @{}

function _Get-ThisScriptDir {
    try {
        if ($PSScriptRoot -and -not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { return $PSScriptRoot }
    } catch { }
    try {
        $p = $MyInvocation.MyCommand.Path
        if ($p -and (Test-Path -LiteralPath $p)) { return (Split-Path -Parent $p) }
    } catch { }
    return (Get-Location).Path
}


function _Add-WatchPhaseMs {
    param(
        [hashtable]$PhaseMs,
        [string]$Name,
        [System.Diagnostics.Stopwatch]$Stopwatch
    )
    if (-not $PhaseMs.ContainsKey($Name)) { $PhaseMs[$Name] = 0 }
    $PhaseMs[$Name] = [int64]$PhaseMs[$Name] + [int64]$Stopwatch.ElapsedMilliseconds
}

function _Get-WatcherQueueRoot {
    param([hashtable]$Config)
    if ($Config -and $Config.ContainsKey('queue') -and $Config.queue) {
        if ($Config.queue.ContainsKey('rootDir') -and $Config.queue.rootDir) { return [string]$Config.queue.rootDir }
        if ($Config.queue.ContainsKey('root') -and $Config.queue.root) { return [string]$Config.queue.root }
        if ($Config.queue.ContainsKey('path') -and $Config.queue.path) { return [string]$Config.queue.path }
    }
    return (Join-Path $repoRoot 'queue')
}

function _Get-WatcherConfigHash {
    param([hashtable]$Config)
    $payload = @{
        filters = if ($Config.ContainsKey('filters')) { $Config.filters } else { $null }
        triggers = if ($Config.ContainsKey('triggers')) { $Config.triggers } else { $null }
    }
    try {
        $json = ($payload | ConvertTo-Json -Depth 80 -Compress)
    } catch {
        $json = [string]$payload
    }
    return Get-Sha256TextHex -Text $json
}

function _Read-LocalWatcherCache {
    param([string]$Path)
    $empty = @{ version = 1; entries = @{} }
    if (-not (Test-Path -LiteralPath $Path)) { return $empty }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $empty }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $entries = @{}
        if ($obj -and $obj.entries) {
            foreach ($prop in @($obj.entries.PSObject.Properties)) {
                $entries[[string]$prop.Name] = $prop.Value
            }
        }
        return @{ version = 1; entries = $entries }
    } catch {
        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_LOCAL_CACHE_READ_FAILED' -Message 'Local watcher cache could not be read; rebuilding.' -Data @{ path = $Path; errorMessage = [string]$_.Exception.Message }
        return $empty
    }
}

function _Write-LocalWatcherCache {
    param(
        [string]$Path,
        [hashtable]$Cache
    )
    try {
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $json = (@{ version = 1; writtenAtUtc = (Get-QCTimestamp); entries = $Cache.entries } | ConvertTo-Json -Depth 80)
        Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
        return $true
    } catch {
        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_LOCAL_CACHE_WRITE_FAILED' -Message 'Local watcher cache could not be written.' -Data @{ path = $Path; errorMessage = [string]$_.Exception.Message }
        return $false
    }
}

function _Get-LocalFileSignature {
    param(
        [System.IO.FileInfo]$FileInfo,
        [string]$NormPath,
        [string]$ConfigHash
    )
    return @{
        path = $NormPath
        lastWriteTicksUtc = [int64]$FileInfo.LastWriteTimeUtc.Ticks
        length = [int64]$FileInfo.Length
        configHash = $ConfigHash
    }
}

function _Test-LocalCacheEntryMatches {
    param(
        [object]$Entry,
        [hashtable]$Signature
    )
    if (-not $Entry) { return $false }
    try {
        return (([int64]$Entry.lastWriteTicksUtc -eq [int64]$Signature.lastWriteTicksUtc) -and
            ([int64]$Entry.length -eq [int64]$Signature.length) -and
            ([string]$Entry.configHash -eq [string]$Signature.configHash))
    } catch { return $false }
}


function _Test-LocalHashEntryMatches {
    param(
        [object]$Entry,
        [hashtable]$Signature
    )
    if (-not $Entry) { return $false }
    try {
        return (([int64]$Entry.lastWriteTicksUtc -eq [int64]$Signature.lastWriteTicksUtc) -and
            ([int64]$Entry.length -eq [int64]$Signature.length) -and
            $Entry.sha256)
    } catch { return $false }
}

function _Set-LocalWatcherCacheEntry {
    param(
        [hashtable]$Cache,
        [string]$Key,
        [hashtable]$Signature,
        [AllowNull()][string]$Sha256 = $null,
        [AllowNull()][string]$Outcome = $null
    )
    $entry = @{
        path = [string]$Signature.path
        lastWriteTicksUtc = [int64]$Signature.lastWriteTicksUtc
        length = [int64]$Signature.length
        configHash = [string]$Signature.configHash
        processedAtUtc = (Get-QCTimestamp)
    }
    if ($Sha256) { $entry.sha256 = $Sha256 }
    if ($Outcome) { $entry.outcome = $Outcome }
    $Cache.entries[$Key] = $entry
}

$scriptDir = _Get-ThisScriptDir
$repoRoot = Split-Path -Parent $scriptDir
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = (Join-Path $repoRoot 'appsettings.json')
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\Core.Hashing.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Filters.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Triggers.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.JobFactory.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force -WarningAction SilentlyContinue
$pwConnPath = (Join-Path $repoRoot 'modules\PW.Connection.psm1')
if (-not (Test-Path -LiteralPath $pwConnPath)) {
    throw "PW.Connection.psm1 not found at expected path: $pwConnPath"
}
Import-Module $pwConnPath -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force -WarningAction SilentlyContinue

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not $config.ContainsKey('dryRun')) { $config['dryRun'] = $false }
if ($DryRun.IsPresent) { $config['dryRun'] = $true }
$isDryRun = [bool]$config['dryRun']

$ignoreSampleEvery = 50
if ($config.ContainsKey('logging') -and $config.logging -and $config.logging.ContainsKey('ignoredSampleEvery') -and $config.logging.ignoredSampleEvery) {
    $ignoreSampleEvery = [int]$config.logging.ignoredSampleEvery
}
if ($ignoreSampleEvery -lt 1) { $ignoreSampleEvery = 1 }

if (-not $config.ContainsKey('watchFolders') -or -not $config.watchFolders) {
    $config['watchFolders'] = @()
}

$watchFolders = @($config.watchFolders | ForEach-Object { ($_ -as [string]).Trim() } | Where-Object { $_ })
$hasPwWatchList = ($config.ContainsKey('projectWise') -and $config.projectWise -and ($config.projectWise.ContainsKey('watchList') -and $config.projectWise.watchList))
if ($watchFolders.Count -eq 0 -and -not $hasPwWatchList) { throw "watchFolders is empty and projectWise.watchList not configured." }

Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_START' -Message 'Watch run started.' -Data @{
    appSettingsPath = $AppSettingsPath
    dryRun = $isDryRun
    watchFolderCount = $watchFolders.Count
    maxFiles = $MaxFiles
    ignoredSampleEvery = $ignoreSampleEvery
}

$accepted = 0
$ignored = 0
$filtered = 0
$matched = 0
$enqueued = 0
$duplicates = 0
$skippedStatusSetCurrent = 0
$errors = 0
$localCacheHits = 0
$localCacheMisses = 0
$localCacheSkips = 0
$hashCacheHits = 0
$hashCacheMisses = 0
$triggerRuleCacheUses = 0
$pwDescriptionLookups = 0
$pwDocEnumerations = 0
$pwFoldersScanned = 0
$dedupeChecks = 0
$dbAuditEventWritesAttempted = 0
$dbAuditEventWritesSucceeded = 0
$dbAuditEventWritesSkipped = 0
$auditPollTelemetry = $null
$watcherRanReconciliationScan = $false

$queueRoot = _Get-WatcherQueueRoot -Config $config

try {
    $startupRes = Invoke-QCQueueStartupCheck -Config $config -ClearWatcherActive
    $startupData = if ($startupRes.IsSuccess) { $startupRes.Data } else { @{} }
    $qStates = $null
    if ($startupData.queueStats -and $startupData.queueStats.states) { $qStates = $startupData.queueStats.states }
    $sqPending = 0; $sqRunning = 0; $sqSucceeded = 0; $sqFailed = 0
    if ($qStates) {
        try { $sqPending = [int]$qStates.pending } catch { }
        try { $sqRunning = [int]$qStates.running } catch { }
        try { $sqSucceeded = [int]$qStates.succeeded } catch { }
        try { $sqFailed = [int]$qStates.failed } catch { }
    }
    $sqRec = $startupData.recovery
    $sqRequeued = 0; $sqFailedRec = 0; $sqOrphans = 0
    if ($sqRec) {
        try { $sqRequeued = [int]$sqRec.recoveredToPending } catch { }
        try { $sqFailedRec = [int]$sqRec.recoveredToFailed } catch { }
        try { $sqOrphans = [int]$sqRec.recoveredOrphan } catch { }
    }
    Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_QUEUE_STARTUP' -Message 'Queue startup check completed.' -Data @{
        queueRoot            = $queueRoot
        pending              = $sqPending
        running              = $sqRunning
        succeeded            = $sqSucceeded
        failed               = $sqFailed
        recoveredToPending   = $sqRequeued
        recoveredToFailed    = $sqFailedRec
        recoveredOrphan      = $sqOrphans
        startupErrors        = @($startupData.errors)
    }
} catch {
    Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_QUEUE_STARTUP_FAILED' -Message 'Queue startup check threw.' -Data @{
        queueRoot = $queueRoot
        error     = [string]$_.Exception.Message
    }
}

$localCachePath = Join-Path (Join-Path $queueRoot '_watcher') 'local-file-cache.json'
$triggerFilterConfigHash = _Get-WatcherConfigHash -Config $config
$localWatcherCache = _Read-LocalWatcherCache -Path $localCachePath
$localWatcherCacheDirty = $false

$orderedTriggerRules = @()
try {
    $orderedTriggerRulesRes = Get-OrderedTriggerRules -Config $config
    if (-not $orderedTriggerRulesRes.IsSuccess) { throw $orderedTriggerRulesRes.Message }
    $orderedTriggerRules = @($orderedTriggerRulesRes.Data.rules)
} catch {
    throw "Failed to load ordered trigger rules: $($_.Exception.Message)"
}

$statusSetRules = @()
try {
    if ($config.ContainsKey('triggers') -and $config.triggers -and $config.triggers.ContainsKey('rules') -and $config.triggers.rules) {
        foreach ($r in @($config.triggers.rules)) {
            $rh = ConvertTo-HashtableDeep -Value $r
            if (-not ($rh -is [hashtable])) { continue }
            if (-not ($rh.ContainsKey('enabled'))) { $rh['enabled'] = $true }
            if (-not [bool]$rh.enabled) { continue }
            if (($rh.ContainsKey('jobType') -and ([string]$rh.jobType) -eq 'STATUS_SET_GEN') -and $rh.ContainsKey('grouping') -and $rh.grouping) {
                $g = ConvertTo-HashtableDeep -Value $rh.grouping
                if ($g -is [hashtable]) {
                    $gEnabled = $false
                    try { $gEnabled = [bool]$g.enabled } catch { $gEnabled = $false }
                    $gBy = $null
                    if ($g.ContainsKey('groupBy') -and $g.groupBy) { $gBy = ([string]$g.groupBy).Trim().ToLowerInvariant() }
                    if ($gEnabled -and $gBy -eq 'folder') {
                        $statusSetRules += $rh
                    }
                }
            }
        }
    }
} catch { }

$statusRuleObj = $null
if ($statusSetRules.Count -gt 0) {
    $statusRuleObj = ($statusSetRules | Sort-Object -Property @{ Expression = { [int]$_.priority }; Descending = $true } | Select-Object -First 1)
}

# ProjectWise watchList processing (STATUS_SET_GEN and/or QC_PREPEND).
# This must run even when STATUS_SET_GEN rules are disabled, because QC_PREPEND can be PW-triggered too.
if ($statusSetRules.Count -ge 0) {
    # ProjectWise sources (watchList) — read-only.
    if ($hasPwWatchList) {
        $pwWatchSw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $pwCfg = ConvertTo-HashtableDeep -Value $config.projectWise
            $ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
            $credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
            $watchList = ConvertTo-HashtableDeep -Value $pwCfg.watchList

            # Re-import here to avoid any odd module/session state where exports are not visible.
            Import-Module $pwConnPath -Force -WarningAction SilentlyContinue | Out-Null
            $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
            if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_CONNECT_START' -Message 'Connecting to ProjectWise.' -Data @{
                datasourceName = $ds
                credentialPath = $credPath
            }
            $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
            if (-not $connRes.IsSuccess) { throw ($connRes.Code + ': ' + $connRes.Message) }
            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_CONNECT_OK' -Message 'Connected to ProjectWise.' -Data @{
                datasourceName = $ds
                userName = if ($credRes.Data -and $credRes.Data.userName) { [string]$credRes.Data.userName } else { '' }
            }

            # One-shot reconciliation: walk every locally-built _StatusSet.pdf
            # and push to PW when the local copy is newer / PW is missing it.
            # Gated by -ReconcileStatusSetsFirst so it runs once per restart,
            # not on every watcher tick (and never blocks normal triggering).
            if ($ReconcileStatusSetsFirst.IsPresent) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_RECONCILE_START' -Message 'Reconciling local status sets to ProjectWise.' -Data @{}
                try {
                    $cb = {
                        param($evt)
                        $level = if ([bool]$evt.isSuccess) { 'Information' } else { 'Warning' }
                        $code  = "WATCH_RECONCILE_$([string]$evt.code -replace '^STATUS_SET_RECONCILE_','')"
                        Write-QCJsonLog -Flush -Level $level -Code $code -Message ([string]$evt.message) -Data @{
                            workspaceDir = [string]$evt.workspaceDir
                            pwFolder     = [string]$evt.pwFolder
                            sheetsFolder = [string]$evt.sheetsFolder
                            outputPdf    = [string]$evt.outputPdf
                            data         = $evt.data
                        }
                    }
                    $rec = Invoke-StatusSetReconcile -Config $config -LogCallback $cb
                    if ($rec.IsSuccess) {
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_RECONCILE_DONE' -Message 'Reconciliation completed.' -Data @{
                            counts   = $rec.Data.counts
                            failures = $rec.Data.failures
                            skipped  = $rec.Data.skipped
                        }
                    } else {
                        Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_RECONCILE_FAILED' -Message ([string]$rec.Message) -Data @{ code = [string]$rec.Code }
                    }
                } catch {
                    Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_RECONCILE_FAILED' -Message ('Reconciliation threw: ' + $_.Exception.Message) -Data @{}
                }
            }

            # --- Hybrid scan decision: audit trail (fast) vs full folder scan (reconciliation) ---
            $auditPollerCfg = $null
            $useAuditScan = $false
            if ($config.ContainsKey('auditPoller') -and $config.auditPoller) {
                $auditPollerCfg = ConvertTo-HashtableDeep -Value $config.auditPoller
                try { $useAuditScan = [bool]$auditPollerCfg.enabled } catch { $useAuditScan = $false }
            }

            $counterPath = Join-Path (Join-Path $queueRoot '_watcher') 'audit-poll-cycle.txt'
            $cycleNum = Get-AuditPollCycleCounter -CounterPath $counterPath
            $reconcileEvery = 20
            if ($auditPollerCfg -and $auditPollerCfg.ContainsKey('reconcileEveryNCycles') -and $auditPollerCfg.reconcileEveryNCycles) {
                try { $reconcileEvery = [int]$auditPollerCfg.reconcileEveryNCycles } catch { $reconcileEvery = 20 }
            }
            $isReconciliationCycle = ($cycleNum -ge $reconcileEvery) -and (($cycleNum % $reconcileEvery) -eq 0)
            # First dashboard tick already runs Invoke-StatusSetReconcile via -ReconcileStatusSetsFirst; skip duplicate full PW walk.
            $runFullScan = ((-not $useAuditScan) -or $isReconciliationCycle) -and (-not $ReconcileStatusSetsFirst.IsPresent)

            if ($useAuditScan -and -not $isReconciliationCycle) {
                # --- AUDIT TRAIL SCAN (primary path) ---
                $auditScanSw = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    $lookbackSeconds = 120
                    if ($auditPollerCfg -and $auditPollerCfg.ContainsKey('lookbackSeconds') -and $auditPollerCfg.lookbackSeconds) {
                        try { $lookbackSeconds = [int]$auditPollerCfg.lookbackSeconds } catch { $lookbackSeconds = 120 }
                    }

                    $watermarkPath = Join-Path (Join-Path $queueRoot '_watcher') 'audit-capture-watermark.txt'
                    $pollWindow = Get-AuditTrailPollWindow -Config $config -WatermarkPath $watermarkPath -LookbackSeconds $lookbackSeconds
                    $since = $pollWindow.since
                    $until = $pollWindow.until

                    $watchRootConfigs = @()
                    if ($watchList -and $watchList.ContainsKey('roots') -and $watchList.roots) {
                        $watchRootConfigs = @($watchList.roots | ForEach-Object { ConvertTo-HashtableDeep -Value $_ })
                    }

                    Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_SCAN_START' -Message 'Audit trail scan starting.' -Data @{
                        since = $since.ToString('yyyy-MM-dd HH:mm:ss')
                        until = $until.ToString('yyyy-MM-dd HH:mm:ss')
                        watermarkBefore = $pollWindow.watermarkBefore
                        isFirstCapture = [bool]$pollWindow.isFirstCapture
                        cycleNum = $cycleNum
                        reconcileEvery = $reconcileEvery
                    }

                    $auditRes = Invoke-AuditTrailScan -Config $config -Since $since -Until $until -WatchRootConfigs $watchRootConfigs
                    if (-not $auditRes.IsSuccess) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SCAN_FAILED' -Message "Audit scan failed: $($auditRes.Message)" -Data @{ code = $auditRes.Code }
                        $fallback = $true
                        if ($auditPollerCfg -and $auditPollerCfg.ContainsKey('fallbackToFullScan')) {
                            try { $fallback = [bool]$auditPollerCfg.fallbackToFullScan } catch { $fallback = $true }
                        }
                        if ($fallback) {
                            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_FALLBACK' -Message 'Falling back to full folder scan.' -Data @{}
                            $runFullScan = $true
                        }
                    } else {
                        $auditData = $auditRes.Data
                        $auditCandidates = @($auditData.candidates)
                        try {
                            if ($auditData.stats) {
                                $dbAuditEventWritesAttempted += ([int]$auditData.stats.dbWrites + [int]$auditData.stats.dbSkipped)
                                $dbAuditEventWritesSucceeded += [int]$auditData.stats.dbWrites
                                $dbAuditEventWritesSkipped += [int]$auditData.stats.dbSkipped
                            }
                        } catch { }
                        $watermarkAfterStr = [string]$auditData.watermarkAfter
                        $capturedThrough = $until
                        try {
                            $parsedWm = [DateTime]::Parse($watermarkAfterStr)
                            if ($parsedWm -gt $capturedThrough) { $capturedThrough = $parsedWm }
                        } catch { }
                        [void](Set-AuditTrailCaptureWatermark -WatermarkPath $watermarkPath -CapturedThrough $capturedThrough)

                        $auditPollTelemetry = @{
                            eventsFetched     = [int]$auditData.stats.totalEvents
                            eventsRelevant    = [int]$auditData.stats.relevantEvents
                            candidatesCreated = [int]$auditCandidates.Count
                            watermarkBefore   = $pollWindow.watermarkBefore
                            watermarkAfter    = $watermarkAfterStr
                            durationMs        = [int]$auditData.durationMs
                        }

                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_SCAN_DONE' -Message 'Audit trail scan completed.' -Data @{
                            totalEvents    = $auditData.stats.totalEvents
                            relevantEvents = $auditData.stats.relevantEvents
                            watchMatches   = $auditData.stats.watchMatches
                            sheetsMatches  = $auditData.stats.sheetsMatches
                            candidates     = $auditCandidates.Count
                            durationMs     = $auditData.durationMs
                            watermarkBefore = $pollWindow.watermarkBefore
                            watermarkAfter = $watermarkAfterStr
                            capturedThrough = $capturedThrough.ToString('yyyy-MM-dd HH:mm:ss')
                        }

                        # Process audit candidates through the existing trigger/job/enqueue pipeline.
                        # Audit-sourced candidates that are in Sheets folders get STATUS_SET_GEN (folder-level).
                        # PDF check-ins with QC_Archivist tag get QC_PREPEND (document-level).
                        $auditFoldersSeen = @{}
                        $auditDescCache = @{}
                        foreach ($ac in $auditCandidates) {
                            try {
                                $fp = [string]$ac.resolvedFolder
                                if ([string]::IsNullOrWhiteSpace($fp)) { continue }

                                # sheet_index: sync ownership from PW when emails/state differ (DOCUMENT_ATTR, etc.)
                                if ($ac.objGuid) {
                                    $acAction = [string]$ac.actionName
                                    $syncOwnership = $acAction -in @('DOCUMENT_ATTR', 'DOCUMENT_STATE')
                                    if ($syncOwnership -or [bool]$ac.isSheetsFolder) {
                                        $acWatchRoot = ''
                                        try { if ($ac.watchRoot) { $acWatchRoot = [string]$ac.watchRoot } } catch { }
                                        Sync-PWSheetIndexOwnership -Config $config `
                                            -DocumentGuid ([string]$ac.objGuid) `
                                            -DocumentName ([string]$ac.itemName) `
                                            -FolderPath $fp `
                                            -IsSheetsFolder ([bool]$ac.isSheetsFolder) `
                                            -WatchRoot $acWatchRoot `
                                            -LastAuditEventAt ([string]$ac.actTime) `
                                            -AuditActionName $acAction
                                    }
                                }

                                # STATUS_SET_GEN: one per unique Sheets folder
                                $acEnableStatusSet = $true
                                try { if ($null -ne $ac.enableStatusSet) { $acEnableStatusSet = [bool]$ac.enableStatusSet } } catch { }
                                if ([bool]$ac.isSheetsFolder -and $acEnableStatusSet -and $statusRuleObj -and -not $auditFoldersSeen.ContainsKey($fp.ToLowerInvariant())) {
                                    $auditFoldersSeen[$fp.ToLowerInvariant()] = $true

                                    $allowRes = Test-QCPathAllowed -CandidatePath $fp -Config $config
                                    if (-not $allowRes.IsSuccess -or -not [bool]$allowRes.Data.allowed) {
                                        $filtered++
                                        continue
                                    }

                                    $candidate = @{
                                        path = $fp
                                        fileName = '_folder_'
                                        description = ''
                                        detectedAtUtc = (Get-QCTimestamp)
                                        sourceFolder = $fp
                                        datasourceName = $ds
                                        groupKey = ('STATUS_SET_GEN|' + $fp).ToLowerInvariant()
                                        triggerSource = 'audit_trail'
                                        file = @{
                                            fullName = $fp
                                            length = 0
                                            lastWriteTimeUtc = (Get-QCTimestamp)
                                        }
                                    }

                                    $jobRes = New-QCJobObject -Candidate $candidate -Rule $statusRuleObj -Config $config
                                    if (-not $jobRes.IsSuccess) { continue }
                                    $job = [hashtable]$jobRes.Data.job

                                    $accepted++
                                    $dedupeChecks++
                                    $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                    if (-not $dupRes.IsSuccess) { continue }
                                    $wouldDedupe = [bool]$dupRes.Data.isDuplicate

                                    Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Audit-sourced STATUS_SET_GEN candidate accepted.' -Data @{
                                        jobId = [string]$job['id']; jobType = [string]$job['type']
                                        sourceFolder = $fp; triggerSource = 'audit_trail'
                                        dryRun = $isDryRun; wouldDedupe = $wouldDedupe
                                    }

                                    if (-not $isDryRun -and -not $wouldDedupe) {
                                        $enqRes = Add-QCQueueJob -Job $job -Config $config
                                        if ($enqRes.IsSuccess) { $enqueued++ }
                                    } elseif ($wouldDedupe) { $duplicates++ }
                                }

                                # QC_PREPEND: evaluate PDFs via trigger rules (description tag, etc.)
                                $acEnableQcPrepend = $true
                                try { if ($null -ne $ac.enableQcPrepend) { $acEnableQcPrepend = [bool]$ac.enableQcPrepend } } catch { }
                                $itemName = [string]$ac.itemName
                                $actionName = [string]$ac.actionName
                                if ($acEnableQcPrepend -and $itemName -match '(?i)\.pdf$' -and $itemName -notmatch '(?i)-qc\.pdf$') {
                                    if (Test-QCIsStatusSetOutputPdfName -FileName $itemName) { continue }
                                    try {
                                        $dd = ''
                                        $descKey = ''
                                        try {
                                            if ($ac.objGuid) { $descKey = ('guid|' + [string]$ac.objGuid).ToLowerInvariant() }
                                            if (-not $descKey) { $descKey = ('path|' + $fp.ToLowerInvariant() + '|' + $itemName.ToLowerInvariant()) }
                                        } catch { $descKey = ('path|' + $fp.ToLowerInvariant() + '|' + $itemName.ToLowerInvariant()) }
                                        if ($descKey -and $auditDescCache.ContainsKey($descKey)) {
                                            $dd = [string]$auditDescCache[$descKey]
                                        } else {
                                            $pwDescriptionLookups++
                                            $dd = Get-PWDocumentDescriptionForFolder -FolderPath $fp -DocumentName $itemName -DocumentGuid ([string]$ac.objGuid)
                                            if ($descKey) { $auditDescCache[$descKey] = [string]$dd }
                                        }
                                        if ($dd.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                                            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (no QC_Archivist in description).' -Data @{
                                                path = ($fp + '\' + $itemName); actionName = $actionName
                                            }
                                            continue
                                        }
                                        $candidate = @{
                                            path            = ($fp + '\' + $itemName)
                                            fileName        = $itemName
                                            description     = [string]$dd
                                            detectedAtUtc   = (Get-QCTimestamp)
                                            sourceFolder    = $fp
                                            datasourceName  = $ds
                                            triggerSource   = 'audit_trail'
                                            auditActionName = $actionName
                                            file            = @{
                                                fullName         = ($fp + '\' + $itemName)
                                                length           = 0
                                                lastWriteTimeUtc = (Get-QCTimestamp)
                                            }
                                        }

                                        $allowRes = Test-QCPathAllowed -CandidatePath ([string]$candidate.path) -Config $config
                                        if (-not $allowRes.IsSuccess -or -not [bool]$allowRes.Data.allowed) {
                                            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped by path filter.' -Data @{
                                                path = [string]$candidate.path; actionName = $actionName
                                            }
                                            $filtered++
                                            continue
                                        }

                                        $matchRes = Test-QCTriggerCandidate -Candidate $candidate -OrderedRules $orderedTriggerRules -Config $config -TriggerType 'pw'
                                        if (-not $matchRes.IsSuccess) {
                                            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Trigger evaluation failed for audit PDF.' -Data @{
                                                path = [string]$candidate.path; actionName = $actionName; error = $matchRes.Message
                                            }
                                            continue
                                        }
                                        if (-not [bool]$matchRes.Data.matched) {
                                            $reason = if ($matchRes.Data.ContainsKey('reason')) { [string]$matchRes.Data.reason } else { 'no_match' }
                                            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF did not match any PW trigger rule.' -Data @{
                                                path = [string]$candidate.path
                                                actionName = $actionName
                                                descriptionPreview = if ($dd.Length -gt 120) { $dd.Substring(0, 120) } else { $dd }
                                                reason = $reason
                                            }
                                            continue
                                        }

                                        $ruleObj = $matchRes.Data.rule
                                        if ([string]$ruleObj.jobType -ne 'QC_PREPEND') { continue }

                                        $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
                                        if (-not $jobRes.IsSuccess) { continue }
                                        $job = [hashtable]$jobRes.Data.job
                                        $accepted++
                                        $dedupeChecks++
                                        $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                        $wouldDedupe = if ($dupRes.IsSuccess) { [bool]$dupRes.Data.isDuplicate } else { $false }

                                        Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Audit-sourced QC_PREPEND candidate accepted.' -Data @{
                                            jobId          = [string]$job['id']
                                            jobType        = 'QC_PREPEND'
                                            sourcePath     = ($fp + '\' + $itemName)
                                            triggerSource  = 'audit_trail'
                                            auditActionName = $actionName
                                            dryRun         = $isDryRun
                                            wouldDedupe    = $wouldDedupe
                                        }

                                        if (-not $isDryRun -and -not $wouldDedupe) {
                                            $enqRes = Add-QCQueueJob -Job $job -Config $config
                                            if ($enqRes.IsSuccess) { $enqueued++ }
                                        } elseif ($wouldDedupe) { $duplicates++ }
                                    } catch {
                                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit QC_PREPEND evaluation threw.' -Data @{
                                            path = ($fp + '\' + $itemName)
                                            actionName = $actionName
                                            error = $_.Exception.Message
                                        }
                                    }
                                }

                                # QC_COMMENT_STATUS_SYNC: *-qc.pdf file updates
                                $acEnableQcCommentSync = $true
                                try { if ($null -ne $ac.enableQcCommentSync) { $acEnableQcCommentSync = [bool]$ac.enableQcCommentSync } } catch { }
                                $commentSyncEnabled = $true
                                if ($config.ContainsKey('qcCommentSync') -and $config.qcCommentSync) {
                                    try { $commentSyncEnabled = [bool]$config.qcCommentSync.enabled } catch { $commentSyncEnabled = $true }
                                }
                                $auditActionsAllowed = @('DOCUMENT_MODIFY', 'DOCUMENT_FILE_REP', 'DOCUMENT_VERSION')
                                if ($config.qcCommentSync -and $config.qcCommentSync.auditActions) {
                                    $auditActionsAllowed = @($config.qcCommentSync.auditActions | ForEach-Object { [string]$_ })
                                }
                                if ($commentSyncEnabled -and $acEnableQcCommentSync -and $itemName -match '(?i)-qc\.pdf$' -and ($auditActionsAllowed -contains $actionName)) {
                                    try {
                                        $pseudoHash = Get-Sha256TextHex -Text (($itemName) + '|' + ([string]$ac.actTime) + '|0|' + $fp)
                                        $candidate = @{
                                            path            = ($fp + '\' + $itemName)
                                            fileName        = $itemName
                                            description     = ''
                                            detectedAtUtc   = (Get-QCTimestamp)
                                            sourceFolder    = $fp
                                            datasourceName  = $ds
                                            documentGuid    = [string]$ac.objGuid
                                            triggerSource   = 'audit_trail'
                                            auditActionName = $actionName
                                            watchRoot       = if ($ac.watchRoot) { [string]$ac.watchRoot } else { '' }
                                            file            = @{
                                                fullName         = ($fp + '\' + $itemName)
                                                length           = 0
                                                lastWriteTimeUtc = [string]$ac.actTime
                                                sha256           = $pseudoHash
                                            }
                                        }

                                        $allowRes = Test-QCPathAllowed -CandidatePath ([string]$candidate.path) -Config $config
                                        if (-not $allowRes.IsSuccess -or -not [bool]$allowRes.Data.allowed) {
                                            $filtered++
                                            continue
                                        }

                                        $matchRes = Test-QCTriggerCandidate -Candidate $candidate -OrderedRules $orderedTriggerRules -Config $config -TriggerType 'pw'
                                        if (-not $matchRes.IsSuccess -or -not [bool]$matchRes.Data.matched) { continue }
                                        $ruleObj = $matchRes.Data.rule
                                        if ([string]$ruleObj.jobType -ne 'QC_COMMENT_STATUS_SYNC') { continue }

                                        $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
                                        if (-not $jobRes.IsSuccess) { continue }
                                        $job = [hashtable]$jobRes.Data.job
                                        $accepted++
                                        $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                        $wouldDedupe = if ($dupRes.IsSuccess) { [bool]$dupRes.Data.isDuplicate } else { $false }

                                        Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Audit-sourced QC_COMMENT_STATUS_SYNC candidate accepted.' -Data @{
                                            jobId           = [string]$job['id']
                                            jobType         = 'QC_COMMENT_STATUS_SYNC'
                                            sourcePath      = ($fp + '\' + $itemName)
                                            triggerSource   = 'audit_trail'
                                            auditActionName = $actionName
                                            dryRun          = $isDryRun
                                            wouldDedupe     = $wouldDedupe
                                        }

                                        if (-not $isDryRun -and -not $wouldDedupe) {
                                            $enqRes = Add-QCQueueJob -Job $job -Config $config
                                            if ($enqRes.IsSuccess) { $enqueued++ }
                                        } elseif ($wouldDedupe) { $duplicates++ }
                                    } catch {
                                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit QC_COMMENT_STATUS_SYNC evaluation threw.' -Data @{
                                            path = ($fp + '\' + $itemName)
                                            actionName = $actionName
                                            error = $_.Exception.Message
                                        }
                                    }
                                }
                            } catch {
                                $errors++
                            }
                        }
                    }
                } catch {
                    Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_AUDIT_SCAN_ERROR' -Message "Audit scan threw: $($_.Exception.Message)" -Data @{ scriptStackTrace = [string]$_.ScriptStackTrace }
                    $errors++
                    $runFullScan = $true
                } finally {
                    $auditScanSw.Stop()
                    _Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'auditTrailScan' -Stopwatch $auditScanSw
                }
            }

            if ($isReconciliationCycle) {
                $watcherRanReconciliationScan = $true
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_RECONCILE_CYCLE' -Message 'Running full folder scan (reconciliation cycle).' -Data @{
                    cycleNum = $cycleNum; reconcileEvery = $reconcileEvery
                }
                Reset-AuditPollCycleCounter -CounterPath $counterPath
            }

            # --- FULL FOLDER SCAN (reconciliation or fallback) ---
            if ($runFullScan) {

            $pwFolders = @()
            if ($watchList -and $watchList.ContainsKey('roots') -and $watchList.roots) {
                foreach ($r in @($watchList.roots)) {
                    $rh = ConvertTo-HashtableDeep -Value $r
                    if (-not ($rh -is [hashtable])) { continue }
                    $rootPath = [string]$rh.path
                    $suffix = if ($rh.ContainsKey('sheetsPathFromProject') -and $rh.sheetsPathFromProject) { [string]$rh.sheetsPathFromProject } else { 'CADD\Sheets' }
                    $projectDepth = 1
                    if ($rh.ContainsKey('projectDepth') -and $null -ne $rh.projectDepth) {
                        try { $projectDepth = [int]$rh.projectDepth } catch { $projectDepth = 1 }
                    }
                    $enableQcPrepend = $false
                    if ($rh.ContainsKey('enableQcPrepend')) { try { $enableQcPrepend = [bool]$rh.enableQcPrepend } catch { $enableQcPrepend = $false } }
                    $enableQcCommentSync = $enableQcPrepend
                    if ($rh.ContainsKey('enableQcCommentSync')) { try { $enableQcCommentSync = [bool]$rh.enableQcCommentSync } catch { } }
                    $enableStatusSet = $false
                    if ($rh.ContainsKey('enableStatusSet')) { try { $enableStatusSet = [bool]$rh.enableStatusSet } catch { $enableStatusSet = $false } }
                    $discovered = @(Find-PWSheetsFoldersUnderRoot -RootPath $rootPath -SheetsSuffix $suffix -DatasourceName $ds -ProjectDepth $projectDepth)
                    foreach ($d in $discovered) {
                        $d['EnableQcPrepend'] = $enableQcPrepend
                        $d['EnableQcCommentSync'] = $enableQcCommentSync
                        $d['EnableStatusSet'] = $enableStatusSet
                        $pwFolders += $d
                    }
                }
            }
            if ($watchList -and $watchList.ContainsKey('folders') -and $watchList.folders) {
                foreach ($f in @($watchList.folders)) {
                    $fh = ConvertTo-HashtableDeep -Value $f
                    if (-not ($fh -is [hashtable])) { continue }
                    $root = [string]$fh.root
                    $path = [string]$fh.path
                    $oneLevelDeep = $false
                    if ($fh.ContainsKey('oneLevelDeep')) { try { $oneLevelDeep = [bool]$fh.oneLevelDeep } catch { $oneLevelDeep = $false } }
                    $enableQcPrepend = $false
                    if ($fh.ContainsKey('enableQcPrepend')) { try { $enableQcPrepend = [bool]$fh.enableQcPrepend } catch { $enableQcPrepend = $false } }
                    $enableQcCommentSync = $enableQcPrepend
                    if ($fh.ContainsKey('enableQcCommentSync')) { try { $enableQcCommentSync = [bool]$fh.enableQcCommentSync } catch { } }
                    $enableStatusSet = $false
                    if ($fh.ContainsKey('enableStatusSet')) { try { $enableStatusSet = [bool]$fh.enableStatusSet } catch { $enableStatusSet = $false } }
                    $full = ($root.TrimEnd('\') + '\' + $path.TrimStart('\')).Trim()
                    $pwFolders += @(@{
                        DatasourceName = $ds
                        FolderPath = $full
                        OneLevelDeep = $oneLevelDeep
                        EnableQcPrepend = $enableQcPrepend
                        EnableQcCommentSync = $enableQcCommentSync
                        EnableStatusSet = $enableStatusSet
                    })
                }
            }

            # Expand oneLevelDeep for explicit folders
            $expanded = @()
            foreach ($e in @($pwFolders)) {
                $expanded += $e
                try {
                    if ($e.OneLevelDeep) {
                        $fp = [string]$e.FolderPath
                        $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS' -Message 'Querying ProjectWise for discipline subfolders under Sheets.' -Data @{
                            folder = $fp
                            inProgress = $true
                        }
                        $kids = @(Get-PWImmediateChildFolders -FolderPath $apiPath)
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS' -Message 'Discipline subfolder listing completed.' -Data @{
                            folder = $fp
                            inProgress = $false
                            childCount = [int]@($kids).Count
                        }
                        if (@($kids).Count -eq 0) {
                            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_NO_CHILDREN' -Message 'oneLevelDeep: no discipline subfolders under this Sheets path; only this folder will be scanned (normal for flat Sheets or empty areas).' -Data @{
                                folder = $fp
                                apiPath = $apiPath
                            }
                        }
                        foreach ($k in $kids) {
                            $kp = Get-PWObjectPropertyValue -Object $k -Name 'FolderPath'
                            if ($kp) {
                                $canonical = ConvertTo-PWCanonicalDocumentsFolderPath -FolderPathProperty ([string]$kp)
                                if (-not $canonical) { continue }
                                $expanded += @{
                                    DatasourceName = $ds
                                    FolderPath = $canonical
                                    OneLevelDeep = $false
                                    EnableQcPrepend = [bool]$e.EnableQcPrepend
                                    EnableStatusSet = [bool]$e.EnableStatusSet
                                }
                            }
                        }
                    }
                } catch {
                    Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_PW_ONELEVEL_EXPAND_FAILED' -Message ('oneLevelDeep expansion failed: ' + $_.Exception.Message) -Data @{
                        folder = [string]$e.FolderPath
                    }
                }
            }
            $pwFolders = $expanded
            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDERS' -Message 'ProjectWise watch folders prepared.' -Data @{
                folderCount = [int]$pwFolders.Count
                sample = @($pwFolders | Select-Object -First 5 | ForEach-Object { [string]$_.FolderPath })
            }

            foreach ($entry in @($pwFolders)) {
                try {
                    $fp = [string]$entry.FolderPath
                    if ([string]::IsNullOrWhiteSpace($fp)) { continue }
                    $pwFoldersScanned++

                    $oneLevelDeep = $false
                    $enableQcPrepend = $false
                    $enableQcCommentSync = $false
                    $enableStatusSet = $false
                    try { $oneLevelDeep = [bool]$entry.OneLevelDeep } catch { $oneLevelDeep = $false }
                    try { $enableQcPrepend = [bool]$entry.EnableQcPrepend } catch { $enableQcPrepend = $false }
                    try { $enableQcCommentSync = [bool]$entry.EnableQcCommentSync } catch { $enableQcCommentSync = $enableQcPrepend }
                    try { $enableStatusSet = [bool]$entry.EnableStatusSet } catch { $enableStatusSet = $false }

                    # Emit a "scan start" event even if filters later skip the folder.
                    Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_SCAN_START' -Message 'PW scanning folder.' -Data @{
                        folder = $fp
                        oneLevelDeep = $oneLevelDeep
                        enableQcPrepend = $enableQcPrepend
                        enableStatusSet = $enableStatusSet
                    }

                    # STATUS_SET_GEN (folder-level)
                    if ($enableStatusSet -and $statusRuleObj) {
                        $allowRes = Test-QCPathAllowed -CandidatePath $fp -Config $config
                        if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
                        if (-not [bool]$allowRes.Data.allowed) {
                            $filtered++
                            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDER_DONE' -Message 'PW folder skipped by filters.' -Data @{
                                folder = $fp
                                reason = 'filtered'
                                enableQcPrepend = $enableQcPrepend
                                enableStatusSet = $enableStatusSet
                            }
                            continue
                        }

                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_START' -Message 'PW status-set folder query started.' -Data @{
                            folder = $fp
                            oneLevelDeep = $oneLevelDeep
                        }
                        $state = Get-StatusSetPWFolderState -FolderPath (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp) -OneLevelDeep:$oneLevelDeep
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_DONE' -Message 'PW status-set folder query completed.' -Data @{
                            folder = $fp
                            oneLevelDeep = $oneLevelDeep
                            pdfCount = [int]$state.pdfCount
                            dgnCount = [int]$state.dgnCount
                            pairedCount = [int]$state.pairedCount
                        }
                        if ([int]$state.pairedCount -gt 0) {
                            # Index paired sheets during reconciliation scan
                            if ($state.pairedSheets -and (Test-QCDatabaseEnabled -Config $config)) {
                                # Email attributes via WithReturnColumns; workflow state via GUID lookup
                                # (WithReturnColumns does not populate WorkflowState in this environment).
                                $nameToAttrs = @{}
                                try {
                                    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp
                                    $searchRows = @(Get-PWDocumentsBySearchWithReturnColumns `
                                        -FolderPath $apiPath -JustThisFolder `
                                        -ColumnsToReturn @('EM_Designer_Email', 'EM_Reviewer_Email') `
                                        -ErrorAction SilentlyContinue)
                                    foreach ($sr in $searchRows) {
                                        $srName = $null
                                        try { $srName = [string]$sr.Name } catch { }
                                        if (-not $srName) { try { $srName = [string]$sr.DocumentName } catch { } }
                                        if (-not $srName) { try { $srName = [string]$sr.FileName } catch { } }
                                        if (-not $srName) { continue }
                                        $de = $null; $re = $null
                                        try {
                                            if ($sr.Attributes) {
                                                foreach ($bag in @($sr.Attributes)) {
                                                    if ($bag -is [System.Collections.IDictionary]) {
                                                        foreach ($k in $bag.Keys) {
                                                            if ([string]$k -eq 'EM_Designer_Email') { $de = [string]$bag[$k] }
                                                            if ([string]$k -eq 'EM_Reviewer_Email') { $re = [string]$bag[$k] }
                                                        }
                                                    }
                                                }
                                            }
                                        } catch { }
                                        $nameToAttrs[$srName.ToLowerInvariant()] = @{
                                            designerEmail = $de; reviewerEmail = $re
                                        }
                                    }
                                } catch { }

                                $stateGuids = @()
                                foreach ($ps in @($state.pairedSheets)) {
                                    if ($ps.pdf -and $ps.pdf.documentGuid) { $stateGuids += [string]$ps.pdf.documentGuid }
                                    if ($ps.dgn -and $ps.dgn.documentGuid) { $stateGuids += [string]$ps.dgn.documentGuid }
                                }
                                $stateByGuid = @{}
                                if ($stateGuids.Count -gt 0) {
                                    try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $stateGuids } catch { }
                                }

                                $sheetIndexRows = @()
                                foreach ($ps in @($state.pairedSheets)) {
                                    try {
                                        $pdfName = if ($ps.pdf -and $ps.pdf.name) { [string]$ps.pdf.name } else { $null }
                                        $dgnName = if ($ps.dgn -and $ps.dgn.name) { [string]$ps.dgn.name } else { $null }
                                        $pdfGuid = if ($ps.pdf -and $ps.pdf.documentGuid) { [string]$ps.pdf.documentGuid } else { $null }
                                        $dgnGuid = if ($ps.dgn -and $ps.dgn.documentGuid) { [string]$ps.dgn.documentGuid } else { $null }

                                        if ($pdfName -and $pdfGuid) {
                                            $a = if ($nameToAttrs.ContainsKey($pdfName.ToLowerInvariant())) { $nameToAttrs[$pdfName.ToLowerInvariant()] } else { @{} }
                                            $pdfState = if ($stateByGuid.ContainsKey($pdfGuid.ToLowerInvariant())) { [string]$stateByGuid[$pdfGuid.ToLowerInvariant()] } else { '' }
                                            $sheetIndexRows += @{
                                                documentGuid = $pdfGuid
                                                documentName = $pdfName
                                                folderPath   = $fp
                                                watchRoot    = [string]$entry.FolderPath
                                                sourceType   = 'pdf'
                                                designerEmail = [string]$a.designerEmail
                                                reviewerEmail = [string]$a.reviewerEmail
                                                pwStateName   = $pdfState
                                            }
                                        }
                                        if ($dgnName -and $dgnGuid) {
                                            $a = if ($nameToAttrs.ContainsKey($dgnName.ToLowerInvariant())) { $nameToAttrs[$dgnName.ToLowerInvariant()] } else { @{} }
                                            $dgnState = if ($stateByGuid.ContainsKey($dgnGuid.ToLowerInvariant())) { [string]$stateByGuid[$dgnGuid.ToLowerInvariant()] } else { '' }
                                            $sheetIndexRows += @{
                                                documentGuid = $dgnGuid
                                                documentName = $dgnName
                                                folderPath   = $fp
                                                watchRoot    = [string]$entry.FolderPath
                                                sourceType   = 'dgn'
                                                designerEmail = [string]$a.designerEmail
                                                reviewerEmail = [string]$a.reviewerEmail
                                                pwStateName   = $dgnState
                                            }
                                        }
                                    } catch { }
                                }

                                if ($sheetIndexRows.Count -gt 0) {
                                    try {
                                        Write-QCSheetIndexBatch -Config $config -Rows $sheetIndexRows
                                    } catch { }
                                }

                                # Link QC PDFs to their source PDF and DGN sheets
                                if ($state.qcPdfDocs) {
                                    foreach ($qc in @($state.qcPdfDocs)) {
                                        try {
                                            $qcStem = [string]$qc.stem
                                            $srcSheet = $state.pairedSheets | Where-Object { $_.stem -eq $qcStem } | Select-Object -First 1
                                            if ($srcSheet) {
                                                if ($srcSheet.pdf -and $srcSheet.pdf.documentGuid) {
                                                    Update-QCSheetQcPdf -Config $config `
                                                        -SourceDocumentGuid ([string]$srcSheet.pdf.documentGuid) `
                                                        -QcPdfGuid ([string]$qc.documentGuid) `
                                                        -QcPdfName ([string]$qc.name)
                                                }
                                                if ($srcSheet.dgn -and $srcSheet.dgn.documentGuid) {
                                                    Update-QCSheetQcPdf -Config $config `
                                                        -SourceDocumentGuid ([string]$srcSheet.dgn.documentGuid) `
                                                        -QcPdfGuid ([string]$qc.documentGuid) `
                                                        -QcPdfName ([string]$qc.name)
                                                }
                                            }
                                        } catch { }
                                    }
                                }
                            }

                            $gateRes = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $fp -FolderState $state
                            $skipUpToDate = ($gateRes.IsSuccess -and -not [bool]$gateRes.Data.shouldEnqueue)
                            if ($skipUpToDate) {
                                $skippedStatusSetCurrent++
                                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SKIP_CURRENT' -Message 'PW folder status set already current; not enqueueing STATUS_SET_GEN.' -Data @{
                                    folder = $fp
                                    pairedCount = [int]$state.pairedCount
                                    gateReason = [string]$gateRes.Data.gateReason
                                    workspaceDir = [string]$gateRes.Data.workspaceDir
                                    manifestPath = [string]$gateRes.Data.manifestPath
                                    compareReasons = if ($gateRes.Data.compare -and $gateRes.Data.compare.reasons) { @($gateRes.Data.compare.reasons) } else { @() }
                                }
                            } else {
                                $candidate = @{
                                    path = $fp
                                    fileName = '_folder_'
                                    description = ''
                                    detectedAtUtc = (Get-QCTimestamp)
                                    sourceFolder = $fp
                                    datasourceName = $ds
                                    groupKey = ('STATUS_SET_GEN|' + $fp).ToLowerInvariant()
                                    folderStateHash = [string]$state.folderStateHash
                                    oneLevelDeep = $oneLevelDeep
                                    statusSet = @{
                                        pairedCount = [int]$state.pairedCount
                                        orderKey = [string]$state.orderKey
                                        pairedSheets = @($state.pairedSheets)
                                    }
                                    file = @{
                                        fullName = $fp
                                        length = 0
                                        lastWriteTimeUtc = (Get-QCTimestamp)
                                    }
                                }

                                $jobRes = New-QCJobObject -Candidate $candidate -Rule $statusRuleObj -Config $config
                                if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
                                $job = [hashtable]$jobRes.Data.job

                                $accepted++
                                $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                                $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                                $wouldEnqueue = (-not $wouldDedupe)
                                $enqueueSkippedReason = $null
                                if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
                                elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

                                Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW folder change candidate accepted (STATUS_SET_GEN).' -Data @{
                                    jobId = [string]$job['id']
                                    jobType = [string]$job['type']
                                    dedupeKey = [string]$job['dedupeKey']
                                    sourcePath = [string]$job['sourcePath']
                                    sourceFolder = [string]$job['sourceFolder']
                                    groupKey = [string]$job['groupKey']
                                    triggeringFile = $fp
                                    ruleId = [string]$job['triggerRule']['id']
                                    dryRun = $isDryRun
                                    wouldEnqueue = $wouldEnqueue
                                    wouldDedupe = $wouldDedupe
                                    enqueueSkippedReason = $enqueueSkippedReason
                                    folderStateHash = [string]$candidate.folderStateHash
                                    pairedCount = [int]$state.pairedCount
                                }

                                if (-not $isDryRun -and -not $wouldDedupe) {
                                    $enqRes = Add-QCQueueJob -Job $job -Config $config
                                    if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                                    $enqueued++
                                } elseif ($wouldDedupe) { $duplicates++ }
                            }
                        } elseif ([int]$state.pdfCount -gt 0 -or [int]$state.dgnCount -gt 0) {
                            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_NO_PAIRS' -Message 'PW folder scanned but no PDF/DGN pairs found.' -Data @{
                                folder = $fp
                                oneLevelDeep = $oneLevelDeep
                                pdfCount = [int]$state.pdfCount
                                dgnCount = [int]$state.dgnCount
                            }
                        }
                    }

                    # QC_PREPEND (description tag)
                    if ([bool]$entry.EnableQcPrepend) {
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_DOC_SCAN_START' -Message 'PW folder doc query started.' -Data @{
                            folder = $fp
                        }
                        $pwDocEnumerations++
                        $docs = Get-PWDocumentsInFolder -FolderPath (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp)
                        $pdfDocs = @()
                        $withDesc = @()
                        $tagged = @()
                        foreach ($d in @($docs)) {
                            $n = Get-PWDocName -Doc $d
                            if (-not $n -or -not ($n -match '(?i)\.pdf$')) { continue }
                            $pdfDocs += $d
                            $dd = Get-PWDocDescription -Doc $d
                            if (-not [string]::IsNullOrWhiteSpace($dd)) {
                                $withDesc += $d
                                if ($dd.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                                    $tagged += $d
                                }
                            }
                        }
                        $sample = @()
                        foreach ($d in @($withDesc | Select-Object -First 2)) {
                            $sn = Get-PWDocName -Doc $d
                            $sd = Get-PWDocDescription -Doc $d
                            if ($sd -and $sd.Length -gt 160) { $sd = $sd.Substring(0, 160) }
                            $sample += (@{ name = $sn; description = $sd })
                        }
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_DOC_SCAN' -Message 'PW folder doc scan completed.' -Data @{
                            folder = $fp
                            docCount = [int](@($docs).Count)
                            pdfCount = [int](@($pdfDocs).Count)
                            withDescriptionCount = [int](@($withDesc).Count)
                            qcArchivistCount = [int](@($tagged).Count)
                            descriptionSample = $sample
                            propertyNamesSample = if (@($docs).Count -gt 0) { @($docs[0].PSObject.Properties | Select-Object -First 30 | ForEach-Object { $_.Name }) } else { @() }
                        }
                        foreach ($doc in @($docs)) {
                            $docName = Get-PWDocName -Doc $doc
                            if (-not $docName -or -not ($docName -match '(?i)\.pdf$')) { continue }
                            if (Test-QCIsStatusSetOutputPdfName -FileName $docName) { continue }
                            $desc = Get-PWDocDescription -Doc $doc
                            if ([string]::IsNullOrWhiteSpace($desc)) { continue }
                            if ($desc.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }

                        Write-QCJsonLog -Level 'Information' -Code 'WATCH_PW_TAGGED' -Message 'PW doc has QC_Archivist tag.' -Data @{
                                folder = $fp
                                fileName = $docName
                                description = $desc
                            }

                            $mod = Get-PWDocLastModifiedUtc -Doc $doc
                            $sz = Get-PWObjectPropertyValue -Object $doc -Name 'FileSize'
                            if (-not $sz) { $sz = Get-PWObjectPropertyValue -Object $doc -Name 'Size' }
                            $pseudo = Get-Sha256TextHex -Text (([string]$docName) + '|' + ([string]$mod) + '|' + ([string]$sz) + '|' + ([string]$fp))

                            $candidate = @{
                                path = ($fp.TrimEnd('\') + '\' + $docName)
                                fileName = $docName
                                description = $desc
                                detectedAtUtc = (Get-QCTimestamp)
                                sourceFolder = $fp
                                datasourceName = $ds
                                file = @{
                                    fullName = ($fp.TrimEnd('\') + '\' + $docName)
                                    length = if ($sz) { [int64]$sz } else { 0 }
                                    lastWriteTimeUtc = $mod
                                    sha256 = $pseudo
                                }
                            }

                            $allowRes = Test-QCPathAllowed -CandidatePath ([string]$candidate.path) -Config $config
                            if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
                            if (-not [bool]$allowRes.Data.allowed) {
                                $filtered++
                                continue
                            }

                            $triggerRuleCacheUses++
                            $matchRes = Test-QCTriggerCandidate -Candidate $candidate -Config $config -OrderedRules $orderedTriggerRules -TriggerType 'pw'
                            if (-not $matchRes.IsSuccess) { throw $matchRes.Message }
                            if (-not [bool]$matchRes.Data.matched) {
                            Write-QCJsonLog -Level 'Information' -Code 'WATCH_PW_NO_MATCH' -Message 'PW doc had QC_Archivist but did not match any PW trigger rule.' -Data @{
                                    path = [string]$candidate.path
                                    fileName = [string]$candidate.fileName
                                    ruleReason = if ($matchRes.Data.ContainsKey('reason')) { [string]$matchRes.Data.reason } else { '' }
                                    candidateDescription = $desc
                                }
                                continue
                            }
                            $ruleObj = $matchRes.Data.rule
                            if ([string]$ruleObj.jobType -ne 'QC_PREPEND') { continue }

                            $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
                            if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
                            $job = [hashtable]$jobRes.Data.job

                            $accepted++
                            $dedupeChecks++
                            $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                            if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                            $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                            $wouldEnqueue = (-not $wouldDedupe)
                            $enqueueSkippedReason = $null
                            if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
                            elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

                            Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW doc accepted (QC_PREPEND via description tag).' -Data @{
                                jobId = [string]$job['id']
                                jobType = [string]$job['type']
                                dedupeKey = [string]$job['dedupeKey']
                                sourcePath = [string]$job['sourcePath']
                                sourceFolder = [string]$job['sourceFolder']
                                triggeringFile = $docName
                                ruleId = [string]$job['triggerRule']['id']
                                dryRun = $isDryRun
                                wouldEnqueue = $wouldEnqueue
                                wouldDedupe = $wouldDedupe
                                enqueueSkippedReason = $enqueueSkippedReason
                            }

                            if (-not $isDryRun -and -not $wouldDedupe) {
                                $enqRes = Add-QCQueueJob -Job $job -Config $config
                                if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                                $enqueued++
                            } elseif ($wouldDedupe) { $duplicates++ }
                        }

                        # QC_COMMENT_STATUS_SYNC: *-qc.pdf (no QC_Archivist tag required)
                        $commentSyncEnabledFs = $true
                        if ($config.ContainsKey('qcCommentSync') -and $config.qcCommentSync) {
                            try { $commentSyncEnabledFs = [bool]$config.qcCommentSync.enabled } catch { $commentSyncEnabledFs = $true }
                        }
                        if ($commentSyncEnabledFs -and $enableQcCommentSync) {
                            foreach ($doc in @($docs)) {
                                $docName = Get-PWDocName -Doc $doc
                                if (-not $docName -or -not ($docName -match '(?i)-qc\.pdf$')) { continue }

                                $mod = Get-PWDocLastModifiedUtc -Doc $doc
                                $sz = Get-PWObjectPropertyValue -Object $doc -Name 'FileSize'
                                if (-not $sz) { $sz = Get-PWObjectPropertyValue -Object $doc -Name 'Size' }
                                $pseudo = Get-Sha256TextHex -Text (([string]$docName) + '|' + ([string]$mod) + '|' + ([string]$sz) + '|' + ([string]$fp))
                                $dg = $null
                                try { $dg = Get-PWObjectPropertyValue -Object $doc -Name 'DocumentGUID' } catch { }
                                if (-not $dg) { try { $dg = Get-PWObjectPropertyValue -Object $doc -Name 'GUID' } catch { } }

                                $candidate = @{
                                    path = ($fp.TrimEnd('\') + '\' + $docName)
                                    fileName = $docName
                                    description = (Get-PWDocDescription -Doc $doc)
                                    detectedAtUtc = (Get-QCTimestamp)
                                    sourceFolder = $fp
                                    datasourceName = $ds
                                    documentGuid = if ($dg) { [string]$dg } else { '' }
                                    triggerSource = 'pw_full_scan'
                                    file = @{
                                        fullName = ($fp.TrimEnd('\') + '\' + $docName)
                                        length = if ($sz) { [int64]$sz } else { 0 }
                                        lastWriteTimeUtc = $mod
                                        sha256 = $pseudo
                                    }
                                }

                                $allowRes = Test-QCPathAllowed -CandidatePath ([string]$candidate.path) -Config $config
                                if (-not $allowRes.IsSuccess -or -not [bool]$allowRes.Data.allowed) { $filtered++; continue }

                                $matchRes = Test-QCTriggerCandidate -Candidate $candidate -Config $config -OrderedRules $orderedTriggerRules -TriggerType 'pw'
                                if (-not $matchRes.IsSuccess -or -not [bool]$matchRes.Data.matched) { continue }
                                $ruleObj = $matchRes.Data.rule
                                if ([string]$ruleObj.jobType -ne 'QC_COMMENT_STATUS_SYNC') { continue }

                                $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
                                if (-not $jobRes.IsSuccess) { continue }
                                $job = [hashtable]$jobRes.Data.job
                                $accepted++
                                $dedupeChecks++
                                $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                                $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                                $enqueueSkippedReason = if ($wouldDedupe) { 'duplicate' } elseif ($isDryRun) { 'dryRun' } else { $null }

                                Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW doc accepted (QC_COMMENT_STATUS_SYNC via -qc.pdf).' -Data @{
                                    jobId = [string]$job['id']
                                    jobType = [string]$job['type']
                                    dedupeKey = [string]$job['dedupeKey']
                                    sourcePath = [string]$job['sourcePath']
                                    ruleId = [string]$job['triggerRule']['id']
                                    dryRun = $isDryRun
                                    wouldDedupe = $wouldDedupe
                                    enqueueSkippedReason = $enqueueSkippedReason
                                }

                                if (-not $isDryRun -and -not $wouldDedupe) {
                                    $enqRes = Add-QCQueueJob -Job $job -Config $config
                                    if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                                    $enqueued++
                                } elseif ($wouldDedupe) { $duplicates++ }
                            }
                        }
                    }
                    Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDER_DONE' -Message 'PW folder processing completed.' -Data @{
                        folder = $fp
                        enableQcPrepend = $enableQcPrepend
                        enableStatusSet = $enableStatusSet
                    }
                } catch {
                    $errors++
                    $ex = $_.Exception
                    Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_PW_FOLDER_ERROR' -Message 'Error processing PW folder for STATUS_SET_GEN.' -Data @{
                        folder = [string]$entry.FolderPath
                        errorMessage = [string]$_.Exception.Message
                        errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
                        scriptStackTrace = [string]$_.ScriptStackTrace
                    }
                }
            }

            } # end if ($runFullScan)

            Disconnect-PW | Out-Null
        } catch {
            $errors++
            Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_PW_ERROR' -Message 'ProjectWise watchList processing failed.' -Data @{ errorMessage = [string]$_.Exception.Message; scriptStackTrace = [string]$_.ScriptStackTrace }
        } finally {
            $pwWatchSw.Stop()
            _Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'projectWiseWatchList' -Stopwatch $pwWatchSw
        }
    }

    $localStatusSetSw = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($folder in $watchFolders) {
        try {
            if (-not (Test-Path -LiteralPath $folder)) { continue }
            $state = Get-StatusSetLocalFolderState -RootFolder $folder
            if ([int]$state.pairedCount -le 0) { continue }

            $normFolderRes = Normalize-QCPath -Path $folder
            if (-not $normFolderRes.IsSuccess) { throw $normFolderRes.Message }
            $normFolder = [string]$normFolderRes.Data.path

            # use the highest-priority rule (lowest priority number wins; our triggers use higher=more important, but
            # this keeps deterministic selection if multiple rules exist)
            # NOTE: Select-Object -First 1 already returns a single hashtable (or $null).
            # Avoid @() which forces object[] and can break downstream Rule property access.
            $ruleObj = ($statusSetRules | Sort-Object -Property priority | Select-Object -First 1)
            $jobType = 'STATUS_SET_GEN'

            if (-not $ruleObj) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_FS_STATUSSET_RULE_MISSING' -Message 'STATUS_SET_GEN rule not found/enabled; skipping folder status-set enqueue.' -Data @{
                    folder = $folder
                    normFolder = $normFolder
                    pairedCount = [int]$state.pairedCount
                }
                continue
            }

            $gateRes = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $normFolder -FolderState $state
            $skipUpToDate = ($gateRes.IsSuccess -and -not [bool]$gateRes.Data.shouldEnqueue)
            if ($skipUpToDate) {
                $skippedStatusSetCurrent++
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_FS_STATUSSET_SKIP_CURRENT' -Message 'Local folder status set already current; not enqueueing STATUS_SET_GEN.' -Data @{
                    folder = $folder
                    normFolder = $normFolder
                    pairedCount = [int]$state.pairedCount
                    gateReason = [string]$gateRes.Data.gateReason
                    workspaceDir = [string]$gateRes.Data.workspaceDir
                }
                continue
            }

            $candidate = @{
                path = $normFolder
                fileName = '_folder_'
                description = ''
                detectedAtUtc = (Get-QCTimestamp)
                sourceFolder = $normFolder
                groupKey = ($jobType + '|' + $normFolder).ToLowerInvariant()
                folderStateHash = [string]$state.folderStateHash
                oneLevelDeep = $false
                statusSet = @{
                    pairedCount = [int]$state.pairedCount
                    orderKey = [string]$state.orderKey
                    pairedSheets = @($state.pairedSheets)
                }
                file = @{
                    fullName = $folder
                    length = 0
                    lastWriteTimeUtc = (Get-QCTimestamp)
                }
            }

            $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
            if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
            $job = [hashtable]$jobRes.Data.job

            $accepted++
            $dedupeChecks++
            $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
            if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
            $wouldDedupe = [bool]$dupRes.Data.isDuplicate
            $wouldEnqueue = (-not $wouldDedupe)
            $enqueueSkippedReason = $null
            if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
            elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

            Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Folder change candidate accepted (STATUS_SET_GEN).' -Data @{
                jobId = [string]$job['id']
                jobType = [string]$job['type']
                dedupeKey = [string]$job['dedupeKey']
                sourcePath = [string]$job['sourcePath']
                sourceFolder = [string]$job['sourceFolder']
                groupKey = [string]$job['groupKey']
                triggeringFile = $folder
                ruleId = [string]$job['triggerRule']['id']
                dryRun = $isDryRun
                wouldEnqueue = $wouldEnqueue
                wouldDedupe = $wouldDedupe
                enqueueSkippedReason = $enqueueSkippedReason
                folderStateHash = [string]$candidate.folderStateHash
                pairedCount = [int]$state.pairedCount
                orderKeyLines = (if ($state.orderKey) { ([string]$state.orderKey -split "`n").Count } else { 0 })
            }

            if (-not $isDryRun -and -not $wouldDedupe) {
                $enqRes = Add-QCQueueJob -Job $job -Config $config
                if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                $enqueued++
            } elseif ($wouldDedupe) {
                $duplicates++
            }
        } catch {
            $errors++
            $ex = $_.Exception
            Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_FOLDER_ERROR' -Message 'Error processing folder for STATUS_SET_GEN.' -Data @{
                folder = $folder
                errorMessage = [string]$_.Exception.Message
                errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
                scriptStackTrace = [string]$_.ScriptStackTrace
            }
        }
    }
    $localStatusSetSw.Stop()
    _Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'localStatusSetFolders' -Stopwatch $localStatusSetSw
    $phaseCounts['statusSetRules'] = [int]$statusSetRules.Count
}

$localDiscoverSw = [System.Diagnostics.Stopwatch]::StartNew()
$fileItems = @()
foreach ($folder in $watchFolders) {
    if (-not (Test-Path -LiteralPath $folder)) {
        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_FOLDER_MISSING' -Message 'Watch folder missing.' -Data @{ folder = $folder }
        continue
    }
    $fileItems += Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue
}

if ($MaxFiles -gt 0) { $fileItems = @($fileItems | Select-Object -First $MaxFiles) }
$localDiscoverSw.Stop()
_Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'localDiscover' -Stopwatch $localDiscoverSw
$phaseCounts['localFilesDiscovered'] = [int]$fileItems.Count
$phaseCounts['watchFolders'] = [int]$watchFolders.Count
$phaseCounts['triggerRules'] = [int]$orderedTriggerRules.Count

$localProcessSw = [System.Diagnostics.Stopwatch]::StartNew()
foreach ($fi in $fileItems) {
    try {
        $pathRes = Normalize-QCPath -Path ([string]$fi.FullName)
        if (-not $pathRes.IsSuccess) { throw $pathRes.Message }
        $normPath = [string]$pathRes.Data.path
        $localCacheKey = $normPath.ToLowerInvariant()
        $localSignature = _Get-LocalFileSignature -FileInfo $fi -NormPath $normPath -ConfigHash $triggerFilterConfigHash
        $localCacheEntry = $null
        if ($localWatcherCache.entries.ContainsKey($localCacheKey)) { $localCacheEntry = $localWatcherCache.entries[$localCacheKey] }
        $localCacheEntryMatches = _Test-LocalCacheEntryMatches -Entry $localCacheEntry -Signature $localSignature
        if ($localCacheEntryMatches) { $localCacheHits++ } else { $localCacheMisses++ }
        if ($localCacheEntryMatches -and -not $isDryRun) {
            $localCacheSkips++
            continue
        }

        $allowRes = Test-QCPathAllowed -CandidatePath $normPath -Config $config
        if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
        if (-not [bool]$allowRes.Data.allowed) {
            $filtered++
            if (-not $isDryRun) {
                _Set-LocalWatcherCacheEntry -Cache $localWatcherCache -Key $localCacheKey -Signature $localSignature -Outcome 'filtered'
                $localWatcherCacheDirty = $true
            }
            continue
        }

        $candidate = @{
            path = $normPath
            fileName = [string]$fi.Name
            description = '' # local filesystem has no PW description; triggers should use filename/path/extension.
            detectedAtUtc = (Get-QCTimestamp)
            file = @{
                fullName = [string]$fi.FullName
                length = [int64]$fi.Length
                lastWriteTimeUtc = $fi.LastWriteTimeUtc.ToString('o')
            }
        }
        $sfRes = Normalize-QCPath -Path ([string]$fi.DirectoryName)
        if (-not $sfRes.IsSuccess) { throw $sfRes.Message }
        $candidate.sourceFolder = [string]$sfRes.Data.path

        $triggerRuleCacheUses++
        $matchRes = Test-QCTriggerCandidate -Candidate $candidate -Config $config -OrderedRules $orderedTriggerRules -TriggerType 'fs'
        if (-not $matchRes.IsSuccess) { throw $matchRes.Message }
        if (-not [bool]$matchRes.Data.matched) {
            $ignored++
            if (($ignored % $ignoreSampleEvery) -eq 0) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_IGNORED_SAMPLE' -Message 'Ignored file (no filesystem trigger match).' -Data @{
                    path = $normPath
                    fileName = $candidate.fileName
                    ignoredCount = $ignored
                }
            }
            if (-not $isDryRun) {
                _Set-LocalWatcherCacheEntry -Cache $localWatcherCache -Key $localCacheKey -Signature $localSignature -Outcome 'ignored'
                $localWatcherCacheDirty = $true
            }
            continue
        }

        $matched++

        $ruleObj = $matchRes.Data.rule
        $jobType = [string]$ruleObj.jobType

        $groupingEnabled = $false
        $groupBy = $null
        $grouping = $null
        if ($ruleObj -and $ruleObj.grouping) { $grouping = ConvertTo-HashtableDeep -Value $ruleObj.grouping }
        if ($grouping -is [hashtable]) {
            try { $groupingEnabled = [bool]$grouping.enabled } catch { $groupingEnabled = $false }
            if ($grouping.ContainsKey('groupBy') -and $grouping.groupBy) { $groupBy = ([string]$grouping.groupBy).Trim().ToLowerInvariant() }
        }

        if (-not $groupingEnabled -or $groupBy -ne 'folder' -or $jobType -ne 'STATUS_SET_GEN') {
            # file-level workflows: compute a stable file hash for dedupe (read-only).
            $cachedSha = $null
            try { if ((_Test-LocalHashEntryMatches -Entry $localCacheEntry -Signature $localSignature)) { $cachedSha = [string]$localCacheEntry.sha256 } } catch { $cachedSha = $null }
            if ($cachedSha) {
                $hashCacheHits++
                $candidate.file.sha256 = $cachedSha
            } else {
                $hashCacheMisses++
                $candidate.file.sha256 = Get-Sha256FileHex -Path ([string]$fi.FullName)
            }
        } else {
            # grouped folder workflow: establish groupKey = jobType + sourceFolder
            $candidate.groupKey = ($jobType + '|' + [string]$candidate.sourceFolder).ToLowerInvariant()
        }

        $jobRes = New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config
        if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
        $job = [hashtable]$jobRes.Data.job

        $accepted++
        $dedupeChecks++
        $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
        if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
        $wouldDedupe = [bool]$dupRes.Data.isDuplicate
        $wouldEnqueue = (-not $wouldDedupe)
        $enqueueSkippedReason = $null
        if ($wouldDedupe) {
            $enqueueSkippedReason = 'duplicate'
        } elseif ($isDryRun) {
            $enqueueSkippedReason = 'dryRun'
        }

        Write-QCJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Trigger matched; job accepted.' -Data @{
            jobId = [string]$job['id']
            jobType = [string]$job['type']
            dedupeKey = [string]$job['dedupeKey']
            sourcePath = [string]$job['sourcePath']
            sourceFolder = [string]$job['sourceFolder']
            groupKey = [string]$job['groupKey']
            triggeringFile = [string]$fi.FullName
            ruleId = [string]$job['triggerRule']['id']
            dryRun = $isDryRun
            wouldEnqueue = $wouldEnqueue
            wouldDedupe = $wouldDedupe
            enqueueSkippedReason = $enqueueSkippedReason
        }

        if ($isDryRun) { continue }

        if ($wouldDedupe) {
            $duplicates++
            _Set-LocalWatcherCacheEntry -Cache $localWatcherCache -Key $localCacheKey -Signature $localSignature -Sha256 ([string]$candidate.file.sha256) -Outcome 'duplicate'
            $localWatcherCacheDirty = $true
            continue
        }

        $enqRes = Add-QCQueueJob -Job $job -Config $config
        if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
        $enqueued++
        _Set-LocalWatcherCacheEntry -Cache $localWatcherCache -Key $localCacheKey -Signature $localSignature -Sha256 ([string]$candidate.file.sha256) -Outcome 'enqueued'
        $localWatcherCacheDirty = $true
    } catch {
        $errors++
        $ex = $_.Exception
        Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_FILE_ERROR' -Message 'Error processing file.' -Data @{
            file = [string]$fi.FullName
            errorMessage = [string]$_.Exception.Message
            errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
            scriptStackTrace = [string]$_.ScriptStackTrace
        }
    }
}
$localProcessSw.Stop()
_Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'localProcess' -Stopwatch $localProcessSw

if ($localWatcherCacheDirty) {
    $cacheWriteSw = [System.Diagnostics.Stopwatch]::StartNew()
    [void](_Write-LocalWatcherCache -Path $localCachePath -Cache $localWatcherCache)
    $cacheWriteSw.Stop()
    _Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'localCacheWrite' -Stopwatch $cacheWriteSw
}

$phaseCounts['localCacheEntries'] = [int]$localWatcherCache.entries.Count
$watchRunSw.Stop()

Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_DONE' -Message 'Watch run completed.' -Data @{
    dryRun = $isDryRun
    elapsedMs = [int64]$watchRunSw.ElapsedMilliseconds
    phaseMs = $phaseMs
    phaseCounts = $phaseCounts
    scanned = $fileItems.Count
    filtered = $filtered
    ignored = $ignored
    matched = $matched
    accepted = $accepted
    duplicates = $duplicates
    skippedStatusSetCurrent = $skippedStatusSetCurrent
    enqueued = $enqueued
    errors = $errors
    localCacheHits = $localCacheHits
    localCacheMisses = $localCacheMisses
    localCacheSkips = $localCacheSkips
    localCacheEntries = [int]$localWatcherCache.entries.Count
    hashCacheHits = $hashCacheHits
    hashCacheMisses = $hashCacheMisses
    triggerRuleCacheUses = $triggerRuleCacheUses
    pwDescriptionLookups = $pwDescriptionLookups
    pwDocEnumerations = $pwDocEnumerations
    pwFoldersScanned = $pwFoldersScanned
    dedupeChecks = $dedupeChecks
    dbAuditEventWritesAttempted = $dbAuditEventWritesAttempted
    dbAuditEventWritesSucceeded = $dbAuditEventWritesSucceeded
    dbAuditEventWritesSkipped = $dbAuditEventWritesSkipped
}

if ($auditPollTelemetry) {
    Write-QCPollRunTelemetry -Config $config `
        -EventsFetched $auditPollTelemetry.eventsFetched `
        -EventsRelevant $auditPollTelemetry.eventsRelevant `
        -CandidatesCreated $auditPollTelemetry.candidatesCreated `
        -JobsEnqueued $enqueued `
        -DurationMs $auditPollTelemetry.durationMs `
        -WatermarkBefore $auditPollTelemetry.watermarkBefore `
        -WatermarkAfter $auditPollTelemetry.watermarkAfter `
        -IsReconciliation:$false
} else {
    Write-QCPollRunTelemetry -Config $config `
        -EventsFetched $fileItems.Count `
        -EventsRelevant $matched `
        -CandidatesCreated $accepted `
        -JobsEnqueued $enqueued `
        -DurationMs ([int]$watchRunSw.ElapsedMilliseconds) `
        -IsReconciliation:$watcherRanReconciliationScan
}

exit 0

