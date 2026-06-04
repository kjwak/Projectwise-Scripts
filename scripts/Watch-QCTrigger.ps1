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
  - Run-once by default; use -Continuous or watcher.continuous to keep one PW session open.
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
    [switch]$ReconcileStatusSetsFirst,

    # Keep process alive: connect to PW once, poll in a loop, disconnect on exit.
    [Parameter(Mandatory = $false)]
    [switch]$Continuous,

    [Parameter(Mandatory = $false)]
    [int]$PollIntervalMs = 0
)

$ErrorActionPreference = 'Stop'
$watchRunSw = [System.Diagnostics.Stopwatch]::StartNew()
$phaseMs = @{}
$phaseCounts = @{}

$watcherPassNumber = $null
$passNumberSource = 'unset'
$runMode = 'manual'
$reconciliationReason = $null
$counterPath = $null
$watcherMode = 'audit_only'
$reconciliationTriggerSource = $null
$downtimeSeconds = 0
$auditGapDetected = $false
$watcherPhase = 'session/connect'
$queueDepthSnapshot = $null


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
        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_LOCAL_CACHE_READ_FAILED' -Message 'Local watcher cache could not be read; rebuilding.' -Data @{ path = $Path; errorMessage = [string]$_.Exception.Message }
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
        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_LOCAL_CACHE_WRITE_FAILED' -Message 'Local watcher cache could not be written.' -Data @{ path = $Path; errorMessage = [string]$_.Exception.Message }
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

function _Set-WatcherAuditPollTelemetryFromFile {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$QueueRoot
    )
    $wmPath = Join-Path (Join-Path $QueueRoot '_watcher') 'audit-capture-watermark.txt'
    $wc = Get-AuditTrailCaptureWatermark -Config $Config -WatermarkPath $wmPath
    $ws = if ($wc) { $wc.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
    $script:auditPollTelemetry = @{
        eventsFetched     = 0
        eventsRelevant    = 0
        candidatesCreated = 0
        watermarkBefore   = $ws
        watermarkAfter    = $ws
        durationMs        = 0
    }
}

function _Reset-WatcherPassState {
    $script:phaseMs = @{}
    $script:phaseCounts = @{}
    $script:watchRunSw = [System.Diagnostics.Stopwatch]::StartNew()
    $script:accepted = 0
    $script:ignored = 0
    $script:filtered = 0
    $script:matched = 0
    $script:enqueued = 0
    $script:duplicates = 0
    $script:skippedStatusSetCurrent = 0
    $script:errors = 0
    $script:localCacheHits = 0
    $script:localCacheMisses = 0
    $script:localCacheSkips = 0
    $script:hashCacheHits = 0
    $script:hashCacheMisses = 0
    $script:triggerRuleCacheUses = 0
    $script:pwDescriptionLookups = 0
    $script:pwDocEnumerations = 0
    $script:pwFoldersScanned = 0
    $script:dedupeChecks = 0
    $script:dbAuditEventWritesAttempted = 0
    $script:dbAuditEventWritesSucceeded = 0
    $script:dbAuditEventWritesSkipped = 0
    $script:auditPollTelemetry = $null
    $script:pollRunWatermarkBefore = $null
    $script:pollRunWatermarkAfter = $null
    $script:watcherRanReconciliationScan = $false
    $script:fullScanScheduleInFlightSlotKey = $null
    $script:watcherPassNumber = $null
    $script:passNumberSource = 'unset'
    $script:runMode = 'manual'
    $script:reconciliationReason = $null
    $script:reconciliationTriggerSource = $null
    $script:downtimeSeconds = 0
    $script:auditGapDetected = $false
    $script:watcherPhase = 'session/connect'
    $script:queueDepthSnapshot = $null
}

$scriptDir = _Get-ThisScriptDir
$repoRoot = Split-Path -Parent $scriptDir
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = (Join-Path $repoRoot 'appsettings.json')
}

$modulesRoot = Join-Path $repoRoot 'modules'
$script:WatchModulesRoot = $modulesRoot

function _Watch-EnsureJsonLog {
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) { return $true }
    $resultsPath = Join-Path $script:WatchModulesRoot 'Core.Results.psm1'
    $runtimePath = Join-Path $script:WatchModulesRoot 'Core.Runtime.psm1'
    if (Test-Path -LiteralPath $resultsPath) {
        Import-Module $resultsPath -Force -WarningAction SilentlyContinue | Out-Null
    }
    if (Test-Path -LiteralPath $runtimePath) {
        Import-Module $runtimePath -Force -WarningAction SilentlyContinue | Out-Null
    }
    return [bool](Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)
}

$script:WatchModuleLoadOrder = @(
    'Core.Results.psm1'
    'Core.Paths.psm1'
    'Core.Runtime.psm1'
    'Core.Hashing.psm1'
    'Core.Database.psm1'
    'QC.Filters.psm1'
    'QC.Triggers.psm1'
    'QC.JobFactory.psm1'
    'QC.Queue.Json.psm1'
    'QC.Rendition.psm1'
    'QC.Processors.psm1'
    'QC.WatcherOrchestration.psm1'
    'QC.StatusSet.psm1'
    'PW.Connection.psm1'
    'PW.Users.psm1'
    'PW.Discovery.psm1'
    'PW.AuditPoller.psm1'
)

# Re-import after PW/QC modules; nested Import-Module -Force drops Core.Database/Core.Hashing session exports.
$script:WatchModuleRestoreOrder = @(
    'Core.Results.psm1'
    'Core.Paths.psm1'
    'Core.Runtime.psm1'
    'Core.Hashing.psm1'
    'Core.Database.psm1'
)

# Commands the watcher calls after nested Import-Module -Force can drop session exports.
$script:WatchRequiredCommands = @(
    'Write-QCJsonLog'
    'Get-QCTimestamp'
    'Get-Sha256TextHex'
    'ConvertTo-HashtableDeep'
    'New-QCSuccessResult'
    'New-QCFailureResult'
    'New-QCErrorResult'
    'Test-QCDatabaseEnabled'
    'Write-QCPollRunTelemetry'
    'Mark-QCAuditEventsProcessed'
    'Write-QCSheetIndex'
    'Write-QCSheetIndexBatch'
    'Get-QCAuditWatermarkAgeSeconds'
    'Get-QCPrependAuditActions'
    'Get-QCInitiatedWorkflowStateName'
    'Test-QCWorkflowStateIsQcInitiated'
    'Get-QCFinalizingWorkflowStateName'
    'Test-QCWorkflowStateIsQcFinalizing'
    'Test-QCIsStatusSetOutputPdfName'
    'Invoke-QCReconcileOutputs'
    'Set-QCFullScanScheduleSlotComplete'
    'Get-OrderedTriggerRules'
    'Test-QCPathAllowed'
    'Test-QCTriggerCandidate'
    'New-QCJobObject'
    'Test-QCDuplicateJob'
    'Add-QCQueueJob'
    'Test-QCStatusSetJobInFlight'
    'Add-QCPrependJobForQcInitiatedStateChange'
    'Add-QCPrependJobForQcFinalizingStateChange'
    'Get-StatusSetPWFolderState'
    'Get-StatusSetLocalFolderState'
    'Test-StatusSetWatcherShouldEnqueue'
    'Invoke-StatusSetReconcile'
    'Get-PWCredentialFromFile'
    'Connect-PW'
    'Get-PWImmediateChildFolders'
    'Get-PWObjectPropertyValue'
    'Sync-PWAssociatedSheetWorkflowState'
    'Sync-PWAssociatedSheetMembersToWorkflowState'
    'Sync-PWSheetIndexOwnership'
    'ConvertTo-PWCmdletFolderPath'
    'Find-PWSheetsFoldersUnderRoot'
    'Get-PWDocumentDescriptionForFolder'
    'Test-PWSheetPdfHasMatchingPair'
    'Test-PWFolderResolvable'
    'Get-PWDocumentWorkflowStateName'
    'Get-PWDocumentsInFolder'
    'Get-PWDocName'
    'Resolve-PWAuditDocumentName'
    'Get-PWDocDescription'
    'Get-PWDocLastModifiedUtc'
    'Get-PWDocumentWorkflowStateMapByGuid'
    'Ensure-PWDiscoveryModuleLoaded'
    'Invoke-AuditTrailScan'
    'Sync-AuditPollerWatchFolderGuidCache'
    'Get-AuditTrailPollWindow'
    'Get-AuditTrailCaptureWatermark'
    'Set-AuditTrailCaptureWatermark'
    'Get-AuditPollCycleCounter'
    'Reset-AuditPollCycleCounter'
)

function _Watch-GetMissingRequiredCommands {
    return @($script:WatchRequiredCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
}

function _Watch-RestoreFoundationModules {
    foreach ($modFile in $script:WatchModuleRestoreOrder) {
        $modPath = Join-Path $script:WatchModulesRoot $modFile
        if (Test-Path -LiteralPath $modPath) {
            Import-Module $modPath -Force -WarningAction SilentlyContinue | Out-Null
        }
    }
}

function _Watch-ReloadWatchModules {
    foreach ($modFile in $script:WatchModuleLoadOrder) {
        $modPath = Join-Path $script:WatchModulesRoot $modFile
        if (Test-Path -LiteralPath $modPath) {
            Import-Module $modPath -Force -WarningAction SilentlyContinue | Out-Null
        }
    }
    if (Get-Command -Name 'Ensure-PWDiscoveryModuleLoaded' -ErrorAction SilentlyContinue) {
        [void](Ensure-PWDiscoveryModuleLoaded)
    }
    _Watch-RestoreFoundationModules
    [void](_Watch-EnsureJsonLog)
}

function _Watch-EnsureAllModuleExports {
    [void](_Watch-EnsureJsonLog)
    for ($pass = 0; $pass -lt 3; $pass++) {
        $missing = @(_Watch-GetMissingRequiredCommands)
        if ($missing.Count -eq 0) { return $true }
        _Watch-ReloadWatchModules
    }
    return (@(_Watch-GetMissingRequiredCommands).Count -eq 0)
}

function _Watch-EnsureStatusSetScanExports {
    return (_Watch-EnsureAllModuleExports)
}

function _Watch-EnsureDiscoveryExports {
    return (_Watch-EnsureAllModuleExports)
}

function _Watch-EnsureDatabaseExports {
    return (_Watch-EnsureAllModuleExports)
}

function _Watch-EnsureResultsExports {
    return (_Watch-EnsureAllModuleExports)
}

function _Watch-EnsureAuditPollerExports {
    return (_Watch-EnsureAllModuleExports)
}

function _Watch-EnsureStatusSetExports {
    return (_Watch-EnsureAllModuleExports)
}

function _Watch-WriteJsonLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Level,
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Data,
        [string]$WorkerLabel = '',
        [switch]$IncludeWorkerPid,
        [switch]$Flush
    )

    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        [void](_Watch-EnsureJsonLog)
    }
    $cmd = Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue
    if ($cmd) {
        $invoke = @{
            Level   = $Level
            Code    = $Code
            Message = $Message
        }
        if ($null -ne $Data) { $invoke['Data'] = $Data }
        if ($WorkerLabel) { $invoke['WorkerLabel'] = $WorkerLabel }
        if ($IncludeWorkerPid) { $invoke['IncludeWorkerPid'] = $true }
        if ($Flush) { $invoke['Flush'] = $true }
        & $cmd @invoke
        return
    }
    if (-not $Data) { $Data = @{} }
    $payload = @{
        ts      = (Get-Date).ToString('o')
        level   = $Level
        code    = $Code
        message = $Message
        data    = $Data
    } | ConvertTo-Json -Depth 20 -Compress
    if ($Flush) {
        [Console]::Out.WriteLine($payload)
        [Console]::Out.Flush()
    } else {
        Write-Host $payload
    }
}

Import-Module (Join-Path $modulesRoot 'Core.Results.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Core.Runtime.psm1') -Force -WarningAction SilentlyContinue
if (-not (_Watch-EnsureJsonLog)) {
    throw "Core.Runtime.psm1 did not load (Write-QCJsonLog missing). Repo root: $repoRoot"
}
Import-Module (Join-Path $repoRoot 'modules\Core.Hashing.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Filters.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Triggers.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.JobFactory.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\PW.Users.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force -WarningAction SilentlyContinue
# QC.AuditTriggers (via PW.Discovery) can reload Core.Database and drop session exports; restore before use.
if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
}
if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
    throw "Core.Database.psm1 did not load; Test-QCDatabaseEnabled is unavailable. Repo: $repoRoot"
}
$pwConnPath = (Join-Path $repoRoot 'modules\PW.Connection.psm1')
if (-not (Test-Path -LiteralPath $pwConnPath)) {
    throw "PW.Connection.psm1 not found at expected path: $pwConnPath"
}
Import-Module $pwConnPath -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.WatcherOrchestration.psm1') -Force -WarningAction SilentlyContinue
_Watch-RestoreFoundationModules
if (-not (_Watch-EnsureAllModuleExports)) {
    $missingAtStart = @(_Watch-GetMissingRequiredCommands)
    throw "Required module exports unavailable after imports ($($missingAtStart -join ', ')). Repo root: $repoRoot"
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (Test-QCDatabaseEnabled -Config $config) {
    try {
        $schemaRes = Initialize-QCDatabaseSchema -Config $config
        if (-not $schemaRes.IsSuccess) {
            _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_DB_SCHEMA_INIT_FAILED' -Message 'Database schema initialization failed; sheet_index writes may fail until schema is upgraded.' -Data @{
                code = [string]$schemaRes.Code
                message = [string]$schemaRes.Message
            }
        } elseif ($schemaRes.Code -in @('DB_SCHEMA_INITIALIZED', 'DB_SCHEMA_UPGRADED')) {
            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_DB_SCHEMA_READY' -Message ([string]$schemaRes.Message) -Data @{
                code = [string]$schemaRes.Code
                version = if ($schemaRes.Data.version) { [string]$schemaRes.Data.version } else { $null }
                previousVersion = if ($schemaRes.Data.previousVersion) { [string]$schemaRes.Data.previousVersion } else { $null }
            }
        }
    } catch {
        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_DB_SCHEMA_INIT_FAILED' -Message ('Database schema initialization threw: ' + $_.Exception.Message) -Data @{}
    }
}

if (-not $config.ContainsKey('dryRun')) { $config['dryRun'] = $false }
if ($DryRun.IsPresent) { $config['dryRun'] = $true }
$isDryRun = [bool]$config['dryRun']
$watcherMode = Get-QCWatcherMode -Config $config -ReconcileStatusSetsFirst:$ReconcileStatusSetsFirst.IsPresent
$reconcileStatusSetsOnStart = Get-QCReconcileStatusSetsOnStart -Config $config
$pollIntervalOverride = if ($PollIntervalMs -gt 0) { [int]$PollIntervalMs } else { $null }
$watcherContinuousSettings = Get-QCWatcherContinuousSettings -Config $config -ContinuousSwitch:$Continuous.IsPresent -PollIntervalMsOverride $pollIntervalOverride
$watcherContinuous = [bool]$watcherContinuousSettings.continuous
$watcherPollSleepMs = [int]$watcherContinuousSettings.pollSleepMs

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

_Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_START' -Message 'Watch run started.' -Data @{
    appSettingsPath = $AppSettingsPath
    dryRun = $isDryRun
    watchFolderCount = $watchFolders.Count
    maxFiles = $MaxFiles
    ignoredSampleEvery = $ignoreSampleEvery
    continuous = $watcherContinuous
    pollSleepMs = $watcherPollSleepMs
}

$queueRoot = _Get-WatcherQueueRoot -Config $config

$script:auditRestartOverlapDone = $false
$script:startupOutputsReconcileDone = $false
try {
    $startupRes = Invoke-QCRecoverQueue -Config $config -ClearWatcherActive
    $startupData = if ($startupRes.IsSuccess) { $startupRes.Data } else { @{} }
    $wmAge = $null
    try {
        $wmPathStartup = Join-Path (Join-Path $queueRoot '_watcher') 'audit-capture-watermark.txt'
        $wmAge = Get-QCAuditWatermarkAgeSeconds -Config $config -WatermarkPath $wmPathStartup
    } catch { }
    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_STARTUP_SEQUENCE' -Message 'RecoverQueue completed (audit reconcile on first PW tick).' -Data @{
        watermarkAgeSeconds = $wmAge
        queueRecovery = $startupData.recovery
    }
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
    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_QUEUE_STARTUP' -Message 'Queue startup check completed.' -Data @{
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
    _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_QUEUE_STARTUP_FAILED' -Message 'Queue startup check threw.' -Data @{
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

$watcherTick = 0
$pwSessionOpen = $false
$script:pwConnectFailureStreak = 0
do {
    $watcherTick++
    _Reset-WatcherPassState
    if ($watcherContinuous) {
        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_TICK_START' -Message 'Watcher poll tick started.' -Data @{
            tick = $watcherTick
            pollSleepMs = $watcherPollSleepMs
            pwSessionOpen = $pwSessionOpen
        }
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
            if (-not (_Watch-EnsureAllModuleExports)) {
                $missingAtTick = @(_Watch-GetMissingRequiredCommands)
                throw ('Required module exports unavailable at tick start: ' + ($missingAtTick -join ', '))
            }
            $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
            if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
            if (-not $pwSessionOpen) {
                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_CONNECT_START' -Message 'Connecting to ProjectWise.' -Data @{
                    datasourceName = $ds
                    credentialPath = $credPath
                    continuous = $watcherContinuous
                    tick = $watcherTick
                }
                $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
                if (-not $connRes.IsSuccess) {
                    $script:pwConnectFailureStreak++
                    throw ($connRes.Code + ': ' + $connRes.Message)
                }
                $pwSessionOpen = $true
                $script:pwConnectFailureStreak = 0
                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_CONNECT_OK' -Message 'Connected to ProjectWise.' -Data @{
                    datasourceName = $ds
                    userName = if ($credRes.Data -and $credRes.Data.userName) { [string]$credRes.Data.userName } else { '' }
                    continuous = $watcherContinuous
                }
            }

            # One-shot reconciliation: walk every locally-built _StatusSet.pdf
            # and push to PW when the local copy is newer / PW is missing it.
            # Gated by reconciliation.reconcileStatusSetsOnStart (appsettings) and
            # -ReconcileStatusSetsFirst (dashboard first pass only).
            $runStatusSetReconcile = $reconcileStatusSetsOnStart -and ($watcherTick -eq 1) -and (
                ($watcherMode -in @('reconciliation','hybrid')) -or $ReconcileStatusSetsFirst.IsPresent
            )
            if ($runStatusSetReconcile) {
                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_RECONCILE_START' -Message 'Reconciling local status sets to ProjectWise.' -Data @{}
                try {
                    $cb = {
                        param($evt)
                        $level = if ([bool]$evt.isSuccess) { 'Information' } else { 'Warning' }
                        $code  = "WATCH_RECONCILE_$([string]$evt.code -replace '^STATUS_SET_RECONCILE_','')"
                        _Watch-WriteJsonLog -Flush -Level $level -Code $code -Message ([string]$evt.message) -Data @{
                            workspaceDir = [string]$evt.workspaceDir
                            pwFolder     = [string]$evt.pwFolder
                            sheetsFolder = [string]$evt.sheetsFolder
                            outputPdf    = [string]$evt.outputPdf
                            data         = $evt.data
                        }
                    }
                    $rec = Invoke-StatusSetReconcile -Config $config -LogCallback $cb
                    if ($rec.IsSuccess) {
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_RECONCILE_DONE' -Message 'Reconciliation completed.' -Data @{
                            counts   = $rec.Data.counts
                            failures = $rec.Data.failures
                            skipped  = $rec.Data.skipped
                        }
                    } else {
                        _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_RECONCILE_FAILED' -Message ([string]$rec.Message) -Data @{ code = [string]$rec.Code }
                    }
                } catch {
                    _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_RECONCILE_FAILED' -Message ('Reconciliation threw: ' + $_.Exception.Message) -Data @{}
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
            $cycleNum = $null
            try {
                # Single increment per tick: pass_number in poll_runs must match cycleNum (was double-called before).
                $cycleNum = [int](Get-AuditPollCycleCounter -CounterPath $counterPath)
                $watcherPassNumber = $cycleNum
                $passNumberSource = 'counter'
            } catch {
                $watcherPassNumber = $null
                $passNumberSource = 'error'
                _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_PASS_COUNTER_READ_FAILED' -Message ([string]$_.Exception.Message) -Data @{ counterPath = $counterPath }
            }
            $fullScanPlan = Get-QCFullFolderScanReconciliationPlan -Config $config -CycleNum $cycleNum -QueueRoot $queueRoot
            $isReconciliationCycle = [bool]$fullScanPlan.due
            $reconcileEvery = if ($null -ne $fullScanPlan.reconcileEvery) { [int]$fullScanPlan.reconcileEvery } else { $null }
            $lastWatermark = $null
            try { if ($auditPollTelemetry -and $auditPollTelemetry.watermarkAfter) { $lastWatermark = [datetime]::Parse([string]$auditPollTelemetry.watermarkAfter).ToUniversalTime() } } catch { }
            $recPlan = Get-QCReconciliationPlan -Config $config -WatcherMode $watcherMode -LastSuccessfulAuditWatermark $lastWatermark
            $reconciliationReason = [string]$recPlan.reason
            $reconciliationTriggerSource = [string]$recPlan.triggerSource
            $downtimeSeconds = [int]$recPlan.downtimeSeconds
            $auditGapDetected = [bool]$recPlan.auditGapDetected
            $runFullScan = [bool]$recPlan.shouldRun

            if (($watcherMode -in @('audit_only','hybrid','recovery')) -and $useAuditScan -and -not $runFullScan) {
                # --- AUDIT TRAIL SCAN (primary path) ---
                $auditScanSw = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    $lookbackSeconds = 120
                    if ($auditPollerCfg -and $auditPollerCfg.ContainsKey('lookbackSeconds') -and $auditPollerCfg.lookbackSeconds) {
                        try { $lookbackSeconds = [int]$auditPollerCfg.lookbackSeconds } catch { $lookbackSeconds = 120 }
                    }

                    $watermarkPath = Join-Path (Join-Path $queueRoot '_watcher') 'audit-capture-watermark.txt'
                    $useRestartOverlap = (-not $script:auditRestartOverlapDone)
                    if ($useRestartOverlap) { $script:auditRestartOverlapDone = $true }
                    $pollWindow = Get-AuditTrailPollWindow -Config $config -WatermarkPath $watermarkPath -LookbackSeconds $lookbackSeconds -UseRestartOverlap:$useRestartOverlap
                    $since = $pollWindow.since
                    $until = $pollWindow.until

                    $watchRootConfigs = @()
                    if ($watchList -and $watchList.ContainsKey('roots') -and $watchList.roots) {
                        $watchRootConfigs = @($watchList.roots | ForEach-Object { ConvertTo-HashtableDeep -Value $_ })
                    }

                    $watermarkAgeSeconds = $null
                    try { $watermarkAgeSeconds = Get-QCAuditWatermarkAgeSeconds -Config $config -WatermarkPath $watermarkPath } catch { }
                    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_SCAN_START' -Message 'Audit trail scan starting.' -Data @{
                        sinceUtc = $pollWindow.sinceUtc
                        untilUtc = $pollWindow.untilUtc
                        sinceDisplay = $pollWindow.sinceDisplay
                        untilDisplay = $pollWindow.untilDisplay
                        watermarkBefore = $pollWindow.watermarkBefore
                        isFirstCapture = [bool]$pollWindow.isFirstCapture
                        overlapSecondsUsed = if ($null -ne $pollWindow.overlapSecondsUsed) { [int]$pollWindow.overlapSecondsUsed } else { 0 }
                        restartOverlapUsed = if ($null -ne $pollWindow.restartOverlapUsed) { [bool]$pollWindow.restartOverlapUsed } else { $false }
                        watermarkAgeSeconds = $watermarkAgeSeconds
                        cycleNum = $cycleNum
                        fullScanScheduleTimes = @($fullScanPlan.scheduledTimes)
                        fullScanScheduleDue = [bool]$fullScanPlan.due
                    }

                    $runMode = 'audit'
                    if (-not (_Watch-EnsureAllModuleExports)) {
                        $missingAudit = @(_Watch-GetMissingRequiredCommands)
                        throw ('Required module exports unavailable before audit trail scan: ' + ($missingAudit -join ', '))
                    }
                    $auditRes = Invoke-AuditTrailScan -Config $config -Since $since -Until $until -WatchRootConfigs $watchRootConfigs
                    if (-not $auditRes.IsSuccess) {
                        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SCAN_FAILED' -Message "Audit scan failed: $($auditRes.Message)" -Data @{ code = $auditRes.Code }
                        $wmAfterFail = $pollWindow.watermarkBefore
                        if ([string]::IsNullOrWhiteSpace($wmAfterFail)) {
                            $wmAfterFail = $pollWindow.untilUtc
                        }
                        $script:auditPollTelemetry = @{
                            eventsFetched     = 0
                            eventsRelevant    = 0
                            candidatesCreated = 0
                            watermarkBefore   = $pollWindow.watermarkBefore
                            watermarkAfter    = $wmAfterFail
                            durationMs        = 0
                        }
                        $fallback = $true
                        if ($auditPollerCfg -and $auditPollerCfg.ContainsKey('fallbackToFullScan')) {
                            try { $fallback = [bool]$auditPollerCfg.fallbackToFullScan } catch { $fallback = $true }
                        }
                        if ($fallback) {
                            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_FALLBACK' -Message 'Falling back to full folder scan.' -Data @{}
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
                        # Advance watermark only when dms_audt rows were ingested this tick (latest o_acttime UTC).
                        $maxPwActTime = $null
                        $maxPwActTimeUtc = $null
                        $totalEventsFetched = 0
                        try {
                            if ($auditData.stats -and $null -ne $auditData.stats.totalEvents) {
                                $totalEventsFetched = [int]$auditData.stats.totalEvents
                            }
                            if ($auditData.stats -and $auditData.stats.maxPwActTime) {
                                $maxPwActTime = [string]$auditData.stats.maxPwActTime
                            }
                            if ($auditData.stats -and $auditData.stats.maxPwActTimeUtc) {
                                $maxPwActTimeUtc = [string]$auditData.stats.maxPwActTimeUtc
                            }
                        } catch { }
                        $watermarkAfterStr = $null
                        $capturedThrough = $null
                        if ($totalEventsFetched -gt 0) {
                            $watermarkAfterStr = if (-not [string]::IsNullOrWhiteSpace($maxPwActTimeUtc)) {
                                $maxPwActTimeUtc
                            } elseif ($auditData.watermarkAfter) {
                                [string]$auditData.watermarkAfter
                            } else {
                                $null
                            }
                            if (-not [string]::IsNullOrWhiteSpace($watermarkAfterStr)) {
                                try {
                                    $capturedThrough = [DateTime]::ParseExact(
                                        $watermarkAfterStr.Trim().TrimEnd('Z'),
                                        'yyyy-MM-dd HH:mm:ss',
                                        $null,
                                        [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
                                    )
                                } catch {
                                    $capturedThrough = $null
                                }
                            }
                            if ($capturedThrough) {
                                [void](Set-AuditTrailCaptureWatermark -WatermarkPath $watermarkPath -CapturedThrough $capturedThrough -Config $config)
                            }
                        }

                        $script:pollRunWatermarkBefore = $pollWindow.watermarkBefore
                        $script:pollRunWatermarkAfter = $watermarkAfterStr

                        $script:auditPollTelemetry = @{
                            eventsFetched     = [int]$auditData.stats.totalEvents
                            eventsRelevant    = [int]$auditData.stats.relevantEvents
                            candidatesCreated = [int]$auditCandidates.Count
                            watermarkBefore   = $pollWindow.watermarkBefore
                            watermarkAfter    = $watermarkAfterStr
                            durationMs        = [int]$auditData.durationMs
                            dbWrites          = [int]$auditData.stats.dbWrites
                            dbSkipped         = [int]$auditData.stats.dbSkipped
                        }

                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_SCAN_DONE' -Message 'Audit trail scan completed.' -Data @{
                            totalEvents    = $auditData.stats.totalEvents
                            relevantEvents = $auditData.stats.relevantEvents
                            watchMatches   = $auditData.stats.watchMatches
                            sheetsMatches  = $auditData.stats.sheetsMatches
                            candidates     = $auditCandidates.Count
                            dbWrites       = [int]$auditData.stats.dbWrites
                            dbSkipped      = [int]$auditData.stats.dbSkipped
                            dbRowsPrepared = [int]$auditData.stats.dbRowsPrepared
                            dbRowsNullGuid = [int]$auditData.stats.dbRowsNullGuid
                            dbLastError    = if ($auditData.stats.dbLastError) { [string]$auditData.stats.dbLastError } else { $null }
                            dbUnprocessedLoaded = [int]$auditData.stats.dbUnprocessedLoaded
                            triggerSource  = if ($auditData.stats.triggerSource) { [string]$auditData.stats.triggerSource } else { $null }
                            guidCacheHits  = [int]$auditData.stats.guidCacheHits
                            guidCacheMisses = [int]$auditData.stats.guidCacheMisses
                            guidResolveSkipped = [int]$auditData.stats.guidResolveSkipped
                            failedGuidCacheHits = [int]$auditData.stats.failedGuidCacheHits
                            foldersResolved = [int]$auditData.stats.foldersResolved
                            pagesFetched   = [int]$auditData.stats.pagesFetched
                            eventsTruncated = [bool]$auditData.stats.eventsTruncated
                            durationMs     = [int]$auditData.durationMs
                            watermarkBefore = $pollWindow.watermarkBefore
                            watermarkAfter = $watermarkAfterStr
                            capturedThrough = if ($capturedThrough) { $capturedThrough.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
                            maxPwActTime = $maxPwActTime
                            auditLogicVersion = if ($auditData.stats.auditLogicVersion) { [string]$auditData.stats.auditLogicVersion } else { $null }
                            candidateSkippedActionCode = if ($null -ne $auditData.stats.candidateSkippedActionCode) { [int]$auditData.stats.candidateSkippedActionCode } else { $null }
                            candidateSkippedNoFolder = if ($null -ne $auditData.stats.candidateSkippedNoFolder) { [int]$auditData.stats.candidateSkippedNoFolder } else { $null }
                            candidateSkippedNoWatchMatch = if ($null -ne $auditData.stats.candidateSkippedNoWatchMatch) { [int]$auditData.stats.candidateSkippedNoWatchMatch } else { $null }
                        }

                        # Process audit candidates through the existing trigger/job/enqueue pipeline.
                        # Audit-sourced candidates that are in Sheets folders get STATUS_SET_GEN (folder-level).
                        # PDF check-ins with QC_Archivist tag get QC_PREPEND (document-level).
                        $auditFoldersSeen = @{}
                        $auditDescCache = @{}
                        $auditSheetPairCache = @{}
                        $auditTriggerEventIds = [System.Collections.Generic.List[long]]::new()
                        $qcPrependAuditActions = @(Get-QCPrependAuditActions -Config $config)
                        foreach ($ac in $auditCandidates) {
                            try {
                                if (-not (_Watch-EnsureAllModuleExports)) {
                                    $missingCand = @(_Watch-GetMissingRequiredCommands)
                                    throw ('Required module exports unavailable before audit candidate: ' + ($missingCand -join ', '))
                                }
                                $fp = [string]$ac.resolvedFolder
                                if ([string]::IsNullOrWhiteSpace($fp)) { continue }

                                $candidateAuditEventId = $null
                                try {
                                    if ($null -ne $ac.auditEventId) {
                                        $aeid = [long]$ac.auditEventId
                                        if ($aeid -gt 0) { $candidateAuditEventId = $aeid }
                                    }
                                } catch { }

                                $itemName = [string]$ac.itemName
                                if ([string]::IsNullOrWhiteSpace($itemName) -and $ac.objGuid `
                                    -and (Get-Command -Name 'Resolve-PWAuditDocumentName' -ErrorAction SilentlyContinue)) {
                                    $resolvedName = Resolve-PWAuditDocumentName -DocumentGuid ([string]$ac.objGuid) `
                                        -FolderPath $fp -Config $config
                                    if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
                                        $itemName = [string]$resolvedName
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NAME_RESOLVED' `
                                            -Message 'Resolved audit document name from sheet_index or PW GUID lookup.' -Data @{
                                            documentGuid = [string]$ac.objGuid
                                            documentName = $itemName
                                            folderPath   = $fp
                                            actionName   = [string]$ac.actionName
                                        }
                                    }
                                }
                                $actionName = [string]$ac.actionName
                                # sheet_index: DOCUMENT_ATTR re-reads EM_* and QC_* from PW.
                                # DOCUMENT_STATE: propagate workflow state to associated DGN / PDF / QC PDF siblings.
                                if ($ac.objGuid) {
                                    $acAction = [string]$ac.actionName
                                    $isDocumentState = $acAction -eq 'DOCUMENT_STATE'
                                    $syncAttributes = $acAction -eq 'DOCUMENT_ATTR'
                                    if ($isDocumentState -or $syncAttributes) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_OWNERSHIP_EVENT' -Message 'Audit attr/state event on watchlist document.' -Data @{
                                            actionName     = $acAction
                                            documentName   = $itemName
                                            folderPath     = $fp
                                            documentGuid   = [string]$ac.objGuid
                                            isSheetsFolder = [bool]$ac.isSheetsFolder
                                            triggerSource  = 'audit_trail'
                                        }
                                    }
                                    $acWatchRoot = ''
                                    try { if ($ac.watchRoot) { $acWatchRoot = [string]$ac.watchRoot } } catch { }
                                    if ($isDocumentState) {
                                        if (-not (_Watch-EnsureAllModuleExports)) {
                                            throw 'Required module exports unavailable before DOCUMENT_STATE sync.'
                                        }
                                        if ([string]::IsNullOrWhiteSpace($itemName)) {
                                            throw ('Document name could not be resolved for DOCUMENT_STATE (documentGuid=' + [string]$ac.objGuid + ').')
                                        }
                                        $skipStateWorkflow = $false
                                        if (Get-Command -Name 'Test-QCShouldSkipAuditWorkflowProcessingForEvent' -ErrorAction SilentlyContinue) {
                                            $skipStateWorkflow = Test-QCShouldSkipAuditWorkflowProcessingForEvent -Config $config -ActTime ([string]$ac.actTime)
                                        }
                                        if ($skipStateWorkflow) {
                                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATE_SKIPPED_GO_LIVE' `
                                                -Message 'Skipped DOCUMENT_STATE workflow sync: audit event is before processingGoLiveUtc.' -Data @{
                                                documentGuid = [string]$ac.objGuid
                                                documentName = $itemName
                                                folderPath     = $fp
                                                actTime        = [string]$ac.actTime
                                                auditEventId   = $candidateAuditEventId
                                            }
                                        } else {
                                            $acUserno = $null
                                            try {
                                                if ($null -ne $ac.userno) { $acUserno = [int]$ac.userno }
                                            } catch { $acUserno = $null }
                                            $acAuditIdSync = $null
                                            try {
                                                if ($null -ne $ac.auditEventId) { $acAuditIdSync = [long]$ac.auditEventId }
                                            } catch { $acAuditIdSync = $null }
                                            Sync-PWAssociatedSheetWorkflowState -Config $config `
                                                -DocumentGuid ([string]$ac.objGuid) `
                                                -DocumentName $itemName `
                                                -FolderPath $fp `
                                                -WatchRoot $acWatchRoot `
                                                -LastAuditEventAt ([string]$ac.actTime) `
                                                -AuditEventId $acAuditIdSync `
                                                -DryRun:$isDryRun `
                                                -ChangedByUser $acUserno
                                        }
                                    } elseif ($syncAttributes -or [bool]$ac.isSheetsFolder) {
                                        if (-not (_Watch-EnsureAllModuleExports)) {
                                            throw 'Required module exports unavailable before DOCUMENT_ATTR sync.'
                                        }
                                        $acUsernoAttr = $null
                                        try {
                                            if ($null -ne $ac.userno) { $acUsernoAttr = [int]$ac.userno }
                                        } catch { $acUsernoAttr = $null }
                                        $acAuditIdAttr = $null
                                        try {
                                            if ($null -ne $ac.auditEventId) { $acAuditIdAttr = [long]$ac.auditEventId }
                                        } catch { $acAuditIdAttr = $null }
                                        Sync-PWSheetIndexOwnership -Config $config `
                                            -DocumentGuid ([string]$ac.objGuid) `
                                            -DocumentName $itemName `
                                            -FolderPath $fp `
                                            -IsSheetsFolder ([bool]$ac.isSheetsFolder) `
                                            -WatchRoot $acWatchRoot `
                                            -LastAuditEventAt ([string]$ac.actTime) `
                                            -AuditEventId $acAuditIdAttr `
                                            -AuditActionName $acAction `
                                            -ChangedByUser $acUsernoAttr
                                    }
                                }

                                # STATUS_SET_GEN: one per unique Sheets folder
                                $acEnableStatusSet = $true
                                try { if ($null -ne $ac.enableStatusSet) { $acEnableStatusSet = [bool]$ac.enableStatusSet } } catch { }
                                if ([bool]$ac.isSheetsFolder -and $acEnableStatusSet -and $statusRuleObj -and -not $auditFoldersSeen.ContainsKey($fp.ToLowerInvariant())) {
                                    $auditFoldersSeen[$fp.ToLowerInvariant()] = $true
                                    if (-not (_Watch-EnsureAllModuleExports)) {
                                        throw 'Required module exports unavailable before audit STATUS_SET_GEN scan.'
                                    }

                                    $allowRes = Test-QCPathAllowed -CandidatePath $fp -Config $config
                                    if (-not $allowRes.IsSuccess -or -not [bool]$allowRes.Data.allowed) {
                                        $filtered++
                                        continue
                                    }

                                    $acOneLevelDeep = $true
                                    try { if ($null -ne $ac.oneLevelDeep) { $acOneLevelDeep = [bool]$ac.oneLevelDeep } } catch { }
                                    $acInFlightRes = Test-QCStatusSetJobInFlight -Config $config -SourceFolder $fp
                                    if ($acInFlightRes.IsSuccess -and [bool]$acInFlightRes.Data.inFlight) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATUSSET_SKIP_IN_FLIGHT' -Message 'Audit STATUS_SET_GEN skipped: job already pending or running for folder.' -Data @{
                                            folder = $fp
                                            jobId = [string]$acInFlightRes.Data.jobId
                                            queueState = [string]$acInFlightRes.Data.queueState
                                        }
                                    } else {
                                    if (-not (_Watch-EnsureStatusSetScanExports)) {
                                        throw 'Required module exports unavailable before audit STATUS_SET_GEN scan.'
                                    }
                                    $acStatusSetScanSw = [System.Diagnostics.Stopwatch]::StartNew()
                                    $acState = Get-StatusSetPWFolderState -FolderPath (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp) -OneLevelDeep:$acOneLevelDeep
                                    $acStatusSetScanSw.Stop()
                                    $acListingMethod = ''
                                    $acDocumentCount = 0
                                    $acDurationMs = [int]$acStatusSetScanSw.ElapsedMilliseconds
                                    try { $acListingMethod = [string]$acState.docListingMethod } catch { }
                                    try { $acDocumentCount = [int]$acState.documentCount } catch { $acDocumentCount = 0 }
                                    try { if ($null -ne $acState.scanDurationMs) { $acDurationMs = [int]$acState.scanDurationMs } } catch { }
                                    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_DONE' -Message 'Audit-sourced PW status-set folder query completed.' -Data @{
                                        folder = $fp
                                        folderPath = $fp
                                        scanReason = 'audit_status_set_candidate'
                                        oneLevelDeep = $acOneLevelDeep
                                        pdfCount = [int]$acState.pdfCount
                                        dgnCount = [int]$acState.dgnCount
                                        pairedCount = [int]$acState.pairedCount
                                        documentCount = $acDocumentCount
                                        listingMethod = $acListingMethod
                                        docListingMethod = $acListingMethod
                                        durationMs = $acDurationMs
                                        scannedFolders = @($acState.scannedFolders)
                                        expandedChildFolders = @($acState.expandedChildFolders)
                                        listingDetails = @($acState.listingDetails)
                                    }
                                    if ([int]$acState.pairedCount -le 0) {
                                        if ([int]$acState.pdfCount -gt 0 -or [int]$acState.dgnCount -gt 0) {
                                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATUSSET_NO_PAIRS' -Message 'Audit STATUS_SET_GEN skipped: no PDF/DGN pairs.' -Data @{
                                                folder = $fp; pdfCount = [int]$acState.pdfCount; dgnCount = [int]$acState.dgnCount; docListingMethod = $acListingMethod
                                            }
                                        } else {
                                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATUSSET_NO_DOCS' -Message 'Audit STATUS_SET_GEN skipped: no sheet docs listed.' -Data @{
                                                folder = $fp; docListingMethod = $acListingMethod
                                            }
                                        }
                                    } else {
                                    $acGateRes = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $fp -FolderState $acState
                                    if ($acGateRes.IsSuccess -and -not [bool]$acGateRes.Data.shouldEnqueue) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATUSSET_SKIP_CURRENT' -Message 'Audit STATUS_SET_GEN skipped: manifest current.' -Data @{
                                            folder = $fp
                                            gateReason = [string]$acGateRes.Data.gateReason
                                            docListingMethod = $acListingMethod
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
                                        folderStateHash = [string]$acState.folderStateHash
                                        oneLevelDeep = $acOneLevelDeep
                                        triggerSource = 'audit_trail'
                                        statusSet = @{
                                            pairedCount = [int]$acState.pairedCount
                                            orderKey = [string]$acState.orderKey
                                        }
                                        file = @{
                                            fullName = $fp
                                            length = 0
                                            lastWriteTimeUtc = (Get-QCTimestamp)
                                        }
                                    }

                                    $jobRes = New-QCJobObject -Candidate $candidate -Rule $statusRuleObj -Config $config
                                    if ($jobRes.IsSuccess) {
                                    $job = [hashtable]$jobRes.Data.job

                                    $accepted++
                                    $dedupeChecks++
                                    $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config
                                    if ($dupRes.IsSuccess) {
                                    $wouldDedupe = [bool]$dupRes.Data.isDuplicate

                                    _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Audit-sourced STATUS_SET_GEN candidate accepted.' -Data @{
                                        jobId = [string]$job['id']; jobType = [string]$job['type']
                                        sourceFolder = $fp; triggerSource = 'audit_trail'
                                        dryRun = $isDryRun; wouldDedupe = $wouldDedupe
                                    }

                                    if (-not $isDryRun -and -not $wouldDedupe) {
                                        $enqRes = Add-QCQueueJob -Job $job -Config $config
                                        if ($enqRes.IsSuccess) { $enqueued++ }
                                    } elseif ($wouldDedupe) { $duplicates++ }
                                    }
                                    }
                                    }
                                    }
                                    }
                                }

                                # QC_PREPEND: paired sheet PDFs — QC Initiated state and/or QC_Archivist description tag.
                                $acEnableQcPrepend = $true
                                try { if ($null -ne $ac.enableQcPrepend) { $acEnableQcPrepend = [bool]$ac.enableQcPrepend } } catch { }
                                if ($acEnableQcPrepend -and $itemName -match '(?i)\.pdf$' -and $itemName -notmatch '(?i)-qc\.pdf$') {
                                    if (Test-QCIsStatusSetOutputPdfName -FileName $itemName) { continue }
                                    if ($qcPrependAuditActions -notcontains $actionName) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (action not configured for QC_PREPEND).' -Data @{
                                            path = ($fp + '\' + $itemName); actionName = $actionName; allowedActions = @($qcPrependAuditActions)
                                        }
                                        continue
                                    }
                                    if (-not [bool]$ac.isSheetsFolder) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (QC_PREPEND is limited to Sheets folders with PDF/DGN pairs).' -Data @{
                                            path = ($fp + '\' + $itemName); actionName = $actionName; isSheetsFolder = $false
                                        }
                                        continue
                                    }
                                    if (-not (_Watch-EnsureDiscoveryExports)) {
                                        _Watch-WriteJsonLog -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (PW.Discovery exports unavailable).' -Data @{
                                            path = ($fp + '\' + $itemName); actionName = $actionName; folderPath = $fp
                                        }
                                        continue
                                    }
                                    if (-not (Get-Command -Name 'Test-PWSheetPdfHasMatchingPair' -ErrorAction SilentlyContinue)) {
                                        _Watch-WriteJsonLog -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (Test-PWSheetPdfHasMatchingPair unavailable).' -Data @{
                                            path = ($fp + '\' + $itemName); actionName = $actionName; folderPath = $fp
                                        }
                                        continue
                                    }
                                    if (-not (Test-PWSheetPdfHasMatchingPair -FolderPath $fp -DocumentName $itemName -PairCache $auditSheetPairCache)) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (no matching DGN pair for sheet stem).' -Data @{
                                            path = ($fp + '\' + $itemName); actionName = $actionName
                                        }
                                        continue
                                    }
                                    if ((Get-Command -Name 'Test-PWFolderResolvable' -ErrorAction SilentlyContinue) `
                                            -and -not (Test-PWFolderResolvable -FolderPath $fp)) {
                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (ProjectWise folder path not resolvable).' -Data @{
                                            path = ($fp + '\' + $itemName); actionName = $actionName; folderPath = $fp
                                        }
                                        continue
                                    }
                                    try {
                                        $pwStateForPrepend = ''
                                        if (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue) {
                                            try {
                                                $pwStateForPrepend = [string](Get-PWDocumentWorkflowStateName -FolderPath $fp -DocumentName $itemName -DocumentGuid ([string]$ac.objGuid))
                                            } catch { }
                                        }
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
                                            if (-not (_Watch-EnsureDiscoveryExports)) {
                                                throw "Get-PWDocumentDescriptionForFolder is unavailable (PW.Discovery failed to load). Repo: $repoRoot"
                                            }
                                            $dd = Get-PWDocumentDescriptionForFolder -FolderPath $fp -DocumentName $itemName -DocumentGuid ([string]$ac.objGuid)
                                            if ($descKey) { $auditDescCache[$descKey] = [string]$dd }
                                        }
                                        if ($dd.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (not QC Initiated and no QC_Archivist in description).' -Data @{
                                                path = ($fp + '\' + $itemName); actionName = $actionName; pwStateName = $pwStateForPrepend
                                            }
                                            continue
                                        }
                                        $initiatedStateName = Get-QCInitiatedWorkflowStateName -Config $config
                                        if (-not [string]::IsNullOrWhiteSpace($initiatedStateName) -and (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) {
                                            $pwState = $pwStateForPrepend
                                            if ([string]::IsNullOrWhiteSpace($pwState)) {
                                                try {
                                                    $pwState = [string](Get-PWDocumentWorkflowStateName -FolderPath $fp -DocumentName $itemName -DocumentGuid ([string]$ac.objGuid))
                                                } catch { }
                                            }
                                            if ([string]::IsNullOrWhiteSpace($pwState) -or ($pwState.Trim().ToLowerInvariant() -ne $initiatedStateName.Trim().ToLowerInvariant())) {
                                                _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped (QC_Archivist tag but workflow state is not QC Initiated).' -Data @{
                                                    path = ($fp + '\' + $itemName); actionName = $actionName
                                                    pwStateName = $pwState; requiredState = $initiatedStateName
                                                }
                                                continue
                                            }
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
                                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF skipped by path filter.' -Data @{
                                                path = [string]$candidate.path; actionName = $actionName
                                            }
                                            $filtered++
                                            continue
                                        }

                                        $matchRes = Test-QCTriggerCandidate -Candidate $candidate -OrderedRules $orderedTriggerRules -Config $config -TriggerType 'pw'
                                        if (-not $matchRes.IsSuccess) {
                                            _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Trigger evaluation failed for audit PDF.' -Data @{
                                                path = [string]$candidate.path; actionName = $actionName; error = $matchRes.Message
                                            }
                                            continue
                                        }
                                        if (-not [bool]$matchRes.Data.matched) {
                                            $reason = if ($matchRes.Data.ContainsKey('reason')) { [string]$matchRes.Data.reason } else { 'no_match' }
                                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit PDF did not match any PW trigger rule.' -Data @{
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

                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Audit-sourced QC_PREPEND candidate accepted.' -Data @{
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
                                        $errMsg = [string]$_.Exception.Message
                                        $logLevel = 'Warning'
                                        $logMsg = 'Audit QC_PREPEND evaluation threw.'
                                        if ($errMsg -match '(?i)StateName.*empty string') {
                                            $logLevel = 'Information'
                                            $logMsg = 'Audit PDF skipped (ProjectWise state lookup failed on this path).'
                                        }
                                        _Watch-WriteJsonLog -Flush -Level $logLevel -Code 'WATCH_AUDIT_SKIPPED' -Message $logMsg -Data @{
                                            path = ($fp + '\' + $itemName)
                                            actionName = $actionName
                                            error = $errMsg
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
                                $auditActionsAllowed = @('DOCUMENT_MODIFY', 'DOCUMENT_FILE_REP', 'DOCUMENT_VERSION', 'DOCUMENT_ATTR', 'DOCUMENT_STATE')
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

                                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Audit-sourced QC_COMMENT_STATUS_SYNC candidate accepted.' -Data @{
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
                                        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_SKIPPED' -Message 'Audit QC_COMMENT_STATUS_SYNC evaluation threw.' -Data @{
                                            path = ($fp + '\' + $itemName)
                                            actionName = $actionName
                                            error = $_.Exception.Message
                                        }
                                    }
                                }

                                if ($null -ne $candidateAuditEventId -and $candidateAuditEventId -gt 0) {
                                    [void]$auditTriggerEventIds.Add([long]$candidateAuditEventId)
                                }
                            } catch {
                                $errors++
                                $errDoc = ''
                                $errAction = ''
                                $errFolder = ''
                                try { $errDoc = [string]$itemName } catch { }
                                if ([string]::IsNullOrWhiteSpace($errDoc)) {
                                    try { $errDoc = [string]$ac.itemName } catch { }
                                }
                                try { $errAction = [string]$ac.actionName } catch { }
                                try { $errFolder = [string]$ac.resolvedFolder } catch { }
                                $errData = @{
                                    documentName = $errDoc
                                    documentGuid = [string]$ac.objGuid
                                    actionName   = $errAction
                                    folderPath   = $errFolder
                                    error        = [string]$_.Exception.Message
                                }
                                $missingAfter = @(_Watch-GetMissingRequiredCommands)
                                if ($missingAfter.Count -gt 0) {
                                    $errData['missingCommands'] = @($missingAfter)
                                }
                                _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_CANDIDATE_ERROR' -Message $_.Exception.Message -Data $errData
                            }
                        }
                        if ($auditTriggerEventIds.Count -gt 0 -and (Get-Command -Name 'Mark-QCAuditEventsProcessed' -ErrorAction SilentlyContinue)) {
                            try {
                                $markRes = Mark-QCAuditEventsProcessed -Config $config -EventIds @($auditTriggerEventIds)
                                if ($markRes.IsSuccess -and $markRes.Data) {
                                    _Watch-WriteJsonLog -Level 'Information' -Code 'AUDIT_EVENTS_MARK_PROCESSED' -Message 'Marked audit_events rows processed after trigger evaluation.' -Data @{
                                        requested = $auditTriggerEventIds.Count
                                        marked = [int]$markRes.Data.marked
                                    }
                                }
                            } catch { }
                        }
                        if (-not $script:startupOutputsReconcileDone) {
                            $script:startupOutputsReconcileDone = $true
                            try {
                                $outRec = Invoke-QCReconcileOutputs -Config $config
                                if ($outRec.IsSuccess) {
                                    _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_RECONCILE_OUTPUTS' -Message 'Startup output reconcile snapshot.' -Data $outRec.Data
                                }
                            } catch { }
                        }
                    }
                } catch {
                    _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_AUDIT_SCAN_ERROR' -Message "Audit scan threw: $($_.Exception.Message)" -Data @{ scriptStackTrace = [string]$_.ScriptStackTrace }
                    $errors++
                    $runFullScan = $true
                } finally {
                    $auditScanSw.Stop()
                    _Add-WatchPhaseMs -PhaseMs $phaseMs -Name 'auditTrailScan' -Stopwatch $auditScanSw
                }
            } elseif ($useAuditScan -and -not $script:auditPollTelemetry) {
                _Set-WatcherAuditPollTelemetryFromFile -Config $config -QueueRoot $queueRoot
            }

            if ($isReconciliationCycle) {
                $watcherRanReconciliationScan = $true
                $runMode = 'reconciliation'
                $reconciliationReason = [string]$fullScanPlan.reason
                if ($fullScanPlan.mode -eq 'schedule' -and $fullScanPlan.slotKey) {
                    $script:fullScanScheduleInFlightSlotKey = [string]$fullScanPlan.slotKey
                }
                $scanMsg = if ($fullScanPlan.mode -eq 'schedule') {
                    "Running full folder scan (scheduled time $($fullScanPlan.scheduledTime))."
                } else {
                    "Running full folder scan (reconciliation cycle $cycleNum, every $reconcileEvery)."
                }
                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_RECONCILE_CYCLE' -Message $scanMsg -Data @{
                    cycleNum = $cycleNum
                    reconcileEvery = $reconcileEvery
                    fullScanMode = [string]$fullScanPlan.mode
                    scheduledTime = [string]$fullScanPlan.scheduledTime
                    scheduledTimes = @($fullScanPlan.scheduledTimes)
                    slotKey = [string]$fullScanPlan.slotKey
                }
                $runFullScan = $true
                if ($fullScanPlan.mode -eq 'cycle') {
                    Reset-AuditPollCycleCounter -CounterPath $counterPath
                }
            }

            # --- FULL FOLDER SCAN (reconciliation or fallback) ---
            if ($runFullScan) {

            if (-not (_Watch-EnsureAllModuleExports)) {
                $missingScan = @(_Watch-GetMissingRequiredCommands)
                throw ('PW.Discovery exports unavailable before full folder scan: ' + ($missingScan -join ', '))
            }

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
                        $d['ScanReason'] = 'full_reconciliation_discovered'
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
                        ScanReason = 'full_reconciliation_configured'
                    })
                }
            }

            # Expand oneLevelDeep once during folder preparation.  The prepared list is
            # concrete for reconciliation: parent entries are kept for direct documents, but
            # their OneLevelDeep flag is suppressed when children are appended so the later
            # status-set scan cannot enumerate those same child folders again.
            $expanded = @()
            foreach ($e in @($pwFolders)) {
                $entryForScan = ConvertTo-HashtableDeep -Value $e
                if (-not ($entryForScan -is [hashtable])) {
                    $entryForScan = @{
                        DatasourceName = $ds
                        FolderPath = [string]$e.FolderPath
                        OneLevelDeep = $false
                        EnableQcPrepend = $false
                        EnableQcCommentSync = $false
                        EnableStatusSet = $false
                    }
                }
                if (-not $entryForScan.ContainsKey('ScanReason')) { $entryForScan['ScanReason'] = 'full_reconciliation_prepared' }
                $expanded += $entryForScan
                try {
                    if ($e.OneLevelDeep) {
                        $fp = [string]$e.FolderPath
                        $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS' -Message 'Querying ProjectWise for discipline subfolders under Sheets.' -Data @{
                            folder = $fp
                            inProgress = $true
                        }
                        $kids = @(Get-PWImmediateChildFolders -FolderPath $apiPath)
                        $entryForScan['OneLevelDeep'] = $false
                        $entryForScan['OneLevelDeepSuppressed'] = $true
                        $entryForScan['ExpandedChildFolderCount'] = [int]@($kids).Count
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_EXPAND_PROGRESS' -Message 'Discipline subfolder listing completed.' -Data @{
                            folder = $fp
                            inProgress = $false
                            childCount = [int]@($kids).Count
                        }
                        if (@($kids).Count -eq 0) {
                            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_NO_CHILDREN' -Message 'oneLevelDeep: no discipline subfolders under this Sheets path; parent will be scanned as an exact folder (normal for flat Sheets or empty areas).' -Data @{
                                folder = $fp
                                apiPath = $apiPath
                            }
                        } else {
                            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_ONELEVEL_SUPPRESSED' -Message 'oneLevelDeep expansion completed; parent will be scanned as an exact folder to avoid duplicate child document enumeration.' -Data @{
                                folder = $fp
                                childCount = [int]@($kids).Count
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
                                    EnableQcCommentSync = [bool]$e.EnableQcCommentSync
                                    EnableStatusSet = [bool]$e.EnableStatusSet
                                    ScanReason = 'full_reconciliation_onelevel_child'
                                    ParentFolderPath = $fp
                                }
                            }
                        }
                    }
                } catch {
                    _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_PW_ONELEVEL_EXPAND_FAILED' -Message ('oneLevelDeep expansion failed: ' + $_.Exception.Message) -Data @{
                        folder = [string]$e.FolderPath
                    }
                }
            }
            $dedupedFolders = @()
            $folderByKey = @{}
            foreach ($prepared in @($expanded)) {
                $preparedPath = ''
                try { $preparedPath = [string]$prepared.FolderPath } catch { }
                if ([string]::IsNullOrWhiteSpace($preparedPath)) { continue }
                $preparedKey = $preparedPath.Trim().TrimEnd('\').ToLowerInvariant()
                if ($folderByKey.ContainsKey($preparedKey)) {
                    $existing = $folderByKey[$preparedKey]
                    try { $existing['EnableQcPrepend'] = ([bool]$existing.EnableQcPrepend -or [bool]$prepared.EnableQcPrepend) } catch { }
                    try { $existing['EnableQcCommentSync'] = ([bool]$existing.EnableQcCommentSync -or [bool]$prepared.EnableQcCommentSync) } catch { }
                    try { $existing['EnableStatusSet'] = ([bool]$existing.EnableStatusSet -or [bool]$prepared.EnableStatusSet) } catch { }
                    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDER_DEDUPED' -Message 'Duplicate prepared folder suppressed before reconciliation scan.' -Data @{
                        folder = $preparedPath
                        FolderPath = $preparedPath
                        existingScanReason = [string]$existing.ScanReason
                        duplicateScanReason = [string]$prepared.ScanReason
                    }
                    continue
                }
                $folderByKey[$preparedKey] = $prepared
                $dedupedFolders += $prepared
            }
            $pwFolders = $dedupedFolders
            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDERS' -Message 'ProjectWise watch folders prepared.' -Data @{
                folderCount = [int]$pwFolders.Count
                sample = @($pwFolders | Select-Object -First 5 | ForEach-Object { [string]$_.FolderPath })
            }

            foreach ($entry in @($pwFolders)) {
                $folderPhase = 'folder_init'
                try {
                    $fp = [string]$entry.FolderPath
                    if ([string]::IsNullOrWhiteSpace($fp)) { continue }
                    $pwFoldersScanned++

                    $oneLevelDeep = $false
                    $enableQcPrepend = $false
                    $enableQcCommentSync = $false
                    $enableStatusSet = $false
                    $scanReason = 'full_reconciliation'
                    $parentFolderPath = ''
                    try { $oneLevelDeep = [bool]$entry.OneLevelDeep } catch { $oneLevelDeep = $false }
                    try { $enableQcPrepend = [bool]$entry.EnableQcPrepend } catch { $enableQcPrepend = $false }
                    try { $enableQcCommentSync = [bool]$entry.EnableQcCommentSync } catch { $enableQcCommentSync = $enableQcPrepend }
                    try { $enableStatusSet = [bool]$entry.EnableStatusSet } catch { $enableStatusSet = $false }
                    try { if ($entry.ScanReason) { $scanReason = [string]$entry.ScanReason } } catch { }
                    try { if ($entry.ParentFolderPath) { $parentFolderPath = [string]$entry.ParentFolderPath } } catch { }

                    # Emit a "scan start" event even if filters later skip the folder.
                    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_SCAN_START' -Message 'PW scanning folder.' -Data @{
                        folder = $fp
                        folderPath = $fp
                        oneLevelDeep = $oneLevelDeep
                        enableQcPrepend = $enableQcPrepend
                        enableStatusSet = $enableStatusSet
                        scanReason = $scanReason
                        parentFolderPath = $parentFolderPath
                    }

                    # STATUS_SET_GEN (folder-level)
                    if ($enableStatusSet -and $statusRuleObj) {
                        $folderPhase = 'statusset_path_allowed'
                        $allowRes = Test-QCPathAllowed -CandidatePath $fp -Config $config
                        if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
                        if (-not [bool]$allowRes.Data.allowed) {
                            $filtered++
                            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDER_DONE' -Message 'PW folder skipped by filters.' -Data @{
                                folder = $fp
                                reason = 'filtered'
                                enableQcPrepend = $enableQcPrepend
                                enableStatusSet = $enableStatusSet
                            }
                            continue
                        }

                        $ssInFlightRes = Test-QCStatusSetJobInFlight -Config $config -SourceFolder $fp
                        if ($ssInFlightRes.IsSuccess -and [bool]$ssInFlightRes.Data.inFlight) {
                            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SKIP_IN_FLIGHT' -Message 'STATUS_SET_GEN already pending or running for folder; skipping PW scan and sheet index.' -Data @{
                                folder = $fp
                                jobId = [string]$ssInFlightRes.Data.jobId
                                queueState = [string]$ssInFlightRes.Data.queueState
                            }
                        } else {
                        if (-not (_Watch-EnsureStatusSetScanExports)) {
                            throw 'Required module exports unavailable before full-scan STATUS_SET_GEN.'
                        }
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_START' -Message 'PW status-set folder query started.' -Data @{
                            folder = $fp
                            oneLevelDeep = $oneLevelDeep
                        }
                        $folderPhase = 'statusset_pw_state'
                        $statusSetScanSw = [System.Diagnostics.Stopwatch]::StartNew()
                        $state = Get-StatusSetPWFolderState -FolderPath (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp) -OneLevelDeep:$oneLevelDeep
                        $statusSetScanSw.Stop()
                        $listingMethod = ''
                        $oneLevelRetry = $false
                        $documentCount = 0
                        $durationMs = [int]$statusSetScanSw.ElapsedMilliseconds
                        try { $listingMethod = [string]$state.docListingMethod } catch { }
                        try { $oneLevelRetry = [bool]$state.oneLevelDeepRetry } catch { }
                        try { $documentCount = [int]$state.documentCount } catch { $documentCount = 0 }
                        try { if ($null -ne $state.scanDurationMs) { $durationMs = [int]$state.scanDurationMs } } catch { }
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_DONE' -Message 'PW status-set folder query completed.' -Data @{
                            folder = $fp
                            folderPath = $fp
                            scanReason = $scanReason
                            parentFolderPath = $parentFolderPath
                            oneLevelDeep = $oneLevelDeep
                            pdfCount = [int]$state.pdfCount
                            dgnCount = [int]$state.dgnCount
                            pairedCount = [int]$state.pairedCount
                            documentCount = $documentCount
                            listingMethod = $listingMethod
                            docListingMethod = $listingMethod
                            durationMs = $durationMs
                            scannedFolders = @($state.scannedFolders)
                            expandedChildFolders = @($state.expandedChildFolders)
                            listingDetails = @($state.listingDetails)
                            oneLevelDeepRetry = $oneLevelRetry
                        }
                        if ([int]$state.pairedCount -gt 0) {
                            $gateRes = Test-StatusSetWatcherShouldEnqueue -Config $config -SourceFolder $fp -FolderState $state
                            $skipUpToDate = ($gateRes.IsSuccess -and -not [bool]$gateRes.Data.shouldEnqueue)
                            if ($skipUpToDate) {
                                $skippedStatusSetCurrent++
                                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SKIP_CURRENT' -Message 'PW folder status set already current; not enqueueing STATUS_SET_GEN.' -Data @{
                                    folder = $fp
                                    pairedCount = [int]$state.pairedCount
                                    gateReason = [string]$gateRes.Data.gateReason
                                    workspaceDir = [string]$gateRes.Data.workspaceDir
                                    manifestPath = [string]$gateRes.Data.manifestPath
                                    compareReasons = if ($gateRes.Data.compare -and $gateRes.Data.compare.reasons) { @($gateRes.Data.compare.reasons) } else { @() }
                                }
                            } else {
                                $folderPhase = 'statusset_enqueue'
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

                                _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW folder change candidate accepted (STATUS_SET_GEN).' -Data @{
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

                            # Sheet index is telemetry only — run after enqueue so large folders do not block STATUS_SET_GEN.
                            $skipSheetIndex = $false
                            $idxInFlightRes = Test-QCStatusSetJobInFlight -Config $config -SourceFolder $fp
                            if ($idxInFlightRes.IsSuccess -and [bool]$idxInFlightRes.Data.inFlight) { $skipSheetIndex = $true }
                            if (-not $skipSheetIndex -and $state.pairedSheets -and (Test-QCDatabaseEnabled -Config $config)) {
                                $folderPhase = 'statusset_sheet_index'
                                $indexSw = [System.Diagnostics.Stopwatch]::StartNew()
                                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_INDEX_START' -Message 'Indexing paired sheets for sheet_index with EM/QC attribute reconciliation.' -Data @{
                                    folder = $fp
                                    pairedCount = [int]$state.pairedCount
                                    reconciliationCycle = [bool]$isReconciliationCycle
                                    scheduledTime = [string]$fullScanPlan.scheduledTime
                                }
                                $stateGuids = @()
                                foreach ($ps in @($state.pairedSheets)) {
                                    if ($ps.pdf -and $ps.pdf.documentGuid) { $stateGuids += [string]$ps.pdf.documentGuid }
                                    if ($ps.dgn -and $ps.dgn.documentGuid) { $stateGuids += [string]$ps.dgn.documentGuid }
                                }
                                $stateByGuid = @{}
                                if ($stateGuids.Count -gt 0) {
                                    try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $stateGuids } catch { }
                                }

                                $sheetIndexRows = @(Build-PWSheetIndexRowsForPairedSheets -Config $config `
                                    -FolderPath $fp `
                                    -WatchRoot ([string]$entry.FolderPath) `
                                    -PairedSheets @($state.pairedSheets) `
                                    -StateByGuid $stateByGuid)

                                $indexRowCount = $sheetIndexRows.Count
                                if ($indexRowCount -gt 0) {
                                    try {
                                        Write-QCSheetIndexBatch -Config $config -Rows @($sheetIndexRows)
                                    } catch { }
                                }

                                if ($state.qcPdfDocs) {
                                    foreach ($qc in @($state.qcPdfDocs)) {
                                        try {
                                            $qcStem = [string]$qc.stem
                                            $srcSheet = $null
                                            foreach ($row in @($state.pairedSheets)) {
                                                $rowStem = $null
                                                if ($row -is [hashtable] -and $row.ContainsKey('stem')) { $rowStem = [string]$row['stem'] }
                                                elseif ($row -and $row.PSObject) { try { $rowStem = [string]$row.stem } catch { } }
                                                if ($rowStem -and $rowStem -eq $qcStem) { $srcSheet = $row; break }
                                            }
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
                                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_INDEX_DONE' -Message 'Sheet index update completed.' -Data @{
                                    folder = $fp
                                    pairedCount = [int]$state.pairedCount
                                    indexRowCount = $indexRowCount
                                    durationMs = [int]$indexSw.ElapsedMilliseconds
                                }
                            }
                        } elseif ([int]$state.pdfCount -gt 0 -or [int]$state.dgnCount -gt 0) {
                            _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_NO_PAIRS' -Message 'PW folder scanned but no PDF/DGN pairs found (PDF-only or missing DGN).' -Data @{
                                folder = $fp
                                oneLevelDeep = $oneLevelDeep
                                pdfCount = [int]$state.pdfCount
                                dgnCount = [int]$state.dgnCount
                                docListingMethod = $listingMethod
                            }
                        } else {
                            _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_PW_STATUSSET_NO_DOCS' -Message 'PW status-set listing found no PDF/DGN in folder; STATUS_SET_GEN not enqueued.' -Data @{
                                folder = $fp
                                oneLevelDeep = $oneLevelDeep
                                docListingMethod = $listingMethod
                                oneLevelDeepRetry = $oneLevelRetry
                            }
                        }
                        } # end: not in-flight STATUS_SET_GEN
                    }

                    # QC_PREPEND (description tag)
                    if ([bool]$entry.EnableQcPrepend) {
                        $folderPhase = 'qc_prepend_doc_scan'
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_DOC_SCAN_START' -Message 'PW folder doc query started.' -Data @{
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
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_DOC_SCAN' -Message 'PW folder doc scan completed.' -Data @{
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

                        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_PW_TAGGED' -Message 'PW doc has QC_Archivist tag.' -Data @{
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
                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_PW_NO_MATCH' -Message 'PW doc had QC_Archivist but did not match any PW trigger rule.' -Data @{
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

                            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW doc accepted (QC_PREPEND via description tag).' -Data @{
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

                                _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW doc accepted (QC_COMMENT_STATUS_SYNC via -qc.pdf).' -Data @{
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
                    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_FOLDER_DONE' -Message 'PW folder processing completed.' -Data @{
                        folder = $fp
                        enableQcPrepend = $enableQcPrepend
                        enableStatusSet = $enableStatusSet
                    }
                } catch {
                    $errors++
                    $ex = $_.Exception
                    _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_PW_FOLDER_ERROR' -Message 'Error processing PW watch folder.' -Data @{
                        folder = [string]$entry.FolderPath
                        phase = [string]$folderPhase
                        enableStatusSet = $enableStatusSet
                        enableQcPrepend = $enableQcPrepend
                        errorMessage = [string]$_.Exception.Message
                        errorType = if ($ex) { [string]$ex.GetType().FullName } else { '' }
                        scriptStackTrace = [string]$_.ScriptStackTrace
                    }
                }
            }

            if ($script:fullScanScheduleInFlightSlotKey) {
                try {
                    $slotDone = Set-QCFullScanScheduleSlotComplete -Config $config -SlotKey $script:fullScanScheduleInFlightSlotKey -QueueRoot $queueRoot
                    _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_FULL_SCAN_SCHEDULE_SLOT_DONE' -Message 'Marked scheduled full-scan slot complete.' -Data @{
                        slotKey = $script:fullScanScheduleInFlightSlotKey
                        persisted = [bool]$slotDone
                    }
                } catch {
                    _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_FULL_SCAN_SCHEDULE_SLOT_FAILED' -Message 'Could not persist full-scan schedule slot completion.' -Data @{
                        slotKey = $script:fullScanScheduleInFlightSlotKey
                        error = [string]$_.Exception.Message
                    }
                }
                $script:fullScanScheduleInFlightSlotKey = $null
            }

            } # end if ($runFullScan)

            if (-not $watcherContinuous) {
                Disconnect-PW | Out-Null
                $pwSessionOpen = $false
            }
        } catch {
            $errors++
            if ($watcherContinuous) {
                if ($pwSessionOpen) {
                    try {
                        Disconnect-PW | Out-Null
                        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_DISCONNECT_ON_ERROR' -Message 'ProjectWise session closed after watch error (will reconnect).' -Data @{ tick = $watcherTick }
                    } catch {
                        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_PW_DISCONNECT_ON_ERROR_FAILED' -Message 'Could not disconnect ProjectWise after watch error.' -Data @{ tick = $watcherTick; error = [string]$_.Exception.Message }
                    }
                }
                $pwSessionOpen = $false
                if ($_.Exception.Message -match 'PW_CONNECT_FAILED') {
                    $script:pwConnectFailureStreak++
                }
            }
            _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_PW_ERROR' -Message 'ProjectWise watchList processing failed.' -Data @{
                errorMessage = [string]$_.Exception.Message
                scriptStackTrace = [string]$_.ScriptStackTrace
                pwConnectFailureStreak = $script:pwConnectFailureStreak
            }
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
                _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_FS_STATUSSET_RULE_MISSING' -Message 'STATUS_SET_GEN rule not found/enabled; skipping folder status-set enqueue.' -Data @{
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
                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_FS_STATUSSET_SKIP_CURRENT' -Message 'Local folder status set already current; not enqueueing STATUS_SET_GEN.' -Data @{
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

            _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Folder change candidate accepted (STATUS_SET_GEN).' -Data @{
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
            _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_FOLDER_ERROR' -Message 'Error processing folder for STATUS_SET_GEN.' -Data @{
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
        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_FOLDER_MISSING' -Message 'Watch folder missing.' -Data @{ folder = $folder }
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
                _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_IGNORED_SAMPLE' -Message 'Ignored file (no filesystem trigger match).' -Data @{
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

        _Watch-WriteJsonLog -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Trigger matched; job accepted.' -Data @{
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
        _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_FILE_ERROR' -Message 'Error processing file.' -Data @{
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

_Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_DONE' -Message 'Watch run completed.' -Data @{
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


$telemetryFailOnWriteError = $false
if ($config.ContainsKey('telemetry') -and $config.telemetry -and $config.telemetry.ContainsKey('failOnWriteError')) {
    try { $telemetryFailOnWriteError = [bool]$config.telemetry.failOnWriteError } catch { $telemetryFailOnWriteError = $false }
}
$runId = [guid]::NewGuid().ToString('N')

$dbWriteSw = [System.Diagnostics.Stopwatch]::StartNew()
$auditDurationSec = if ($phaseMs.ContainsKey('auditTrailScan')) { [math]::Round(([decimal]$phaseMs['auditTrailScan']/1000),3) } else { $null }
$reconDurationSec = if ($phaseMs.ContainsKey('fullPwScan')) { [math]::Round(([decimal]$phaseMs['fullPwScan']/1000),3) } else { $null }
$triggerEvalSec = if ($phaseMs.ContainsKey('localProcess')) { [math]::Round(([decimal]$phaseMs['localProcess']/1000),3) } else { $null }
$dedupeSec = if ($phaseMs.ContainsKey('dedupeChecks')) { [math]::Round(([decimal]$phaseMs['dedupeChecks']/1000),3) } else { $null }
$queueWriteSec = if ($phaseMs.ContainsKey('queueWrite')) { [math]::Round(([decimal]$phaseMs['queueWrite']/1000),3) } else { $null }
$cleanupSec = if ($phaseMs.ContainsKey('localCacheWrite')) { [math]::Round(([decimal]$phaseMs['localCacheWrite']/1000),3) } else { $null }
$sleepThrottleSec = if ($phaseMs.ContainsKey('sleepThrottle')) { [math]::Round(([decimal]$phaseMs['sleepThrottle']/1000),3) } else { 0 }
if (-not (_Watch-EnsureAllModuleExports) -or -not (Get-Command -Name 'Write-QCPollRunTelemetry' -ErrorAction SilentlyContinue)) {
    _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_TELEMETRY_SKIPPED' -Message 'Write-QCPollRunTelemetry unavailable; poll telemetry not written.' -Data @{
        runId = $runId
        missingCommands = @(_Watch-GetMissingRequiredCommands)
    }
    $telemetryRes = New-QCFailureResult -Code 'WATCH_TELEMETRY_SKIPPED' -Message 'Write-QCPollRunTelemetry unavailable.' -Data @{}
} else {
    $telemetryRes = Write-QCPollRunTelemetry -Config $config `
    -EventsFetched $(if($script:auditPollTelemetry){$script:auditPollTelemetry.eventsFetched}else{$fileItems.Count}) `
    -EventsRelevant $(if($script:auditPollTelemetry){$script:auditPollTelemetry.eventsRelevant}else{$matched}) `
    -CandidatesCreated $(if($script:auditPollTelemetry){$script:auditPollTelemetry.candidatesCreated}else{$accepted}) `
    -JobsEnqueued $enqueued `
    -DurationMs ([int]$watchRunSw.ElapsedMilliseconds) `
    -WatermarkBefore $(if($script:pollRunWatermarkAfter -or $script:pollRunWatermarkBefore){$script:pollRunWatermarkBefore}else{if($script:auditPollTelemetry){$script:auditPollTelemetry.watermarkBefore}else{$null}}) `
    -WatermarkAfter $(if($script:pollRunWatermarkAfter){$script:pollRunWatermarkAfter}else{if($script:auditPollTelemetry){$script:auditPollTelemetry.watermarkAfter}else{$null}}) `
    -IsReconciliation:$watcherRanReconciliationScan `
    -PassNumber $watcherPassNumber `
    -RunMode $runMode `
    -RunStatus $(if($errors -gt 0){'failed'}else{'succeeded'}) `
    -TotalDurationSeconds ([math]::Round(([decimal]$watchRunSw.ElapsedMilliseconds/1000),3)) `
    -AuditQueryDurationSeconds $auditDurationSec `
    -ReconciliationDurationSeconds $reconDurationSec `
    -TriggerEvalDurationSeconds $triggerEvalSec `
    -DedupeDurationSeconds $dedupeSec `
    -QueueWriteDurationSeconds $queueWriteSec `
    -CleanupDurationSeconds $cleanupSec `
    -SleepThrottleDurationSeconds $sleepThrottleSec `
    -CandidateDocumentsEvaluated $accepted `
    -TriggerMatches $matched `
    -JobsSkippedDedupe $duplicates `
    -WarningCount 0 `
    -ErrorCount $errors `
    -ReconciliationReason $reconciliationReason `
    -ReconciliationTriggerSource $reconciliationTriggerSource `
    -DowntimeSeconds $downtimeSeconds `
    -AuditGapDetected:$auditGapDetected `
    -WatcherPhase 'telemetry_publish' `
    -ThrottleWaitSeconds $sleepThrottleSec `
    -QueueDepthSnapshot $queueDepthSnapshot `
    -RunId $runId `
    -PassNumberSource $passNumberSource
}
$dbWriteSw.Stop()
if (-not $telemetryRes.IsSuccess) {
    _Watch-WriteJsonLog -Flush -Level 'Error' -Code 'WATCH_TELEMETRY_WRITE_FAILED' -Message $telemetryRes.Message -Data @{ runId=$runId; passNumber=$watcherPassNumber; watcherName='qc_watcher' }
    if ($telemetryFailOnWriteError) { throw ('Telemetry write failed: ' + $telemetryRes.Message) }
}

    if ($watcherContinuous) {
        $sleepMs = $watcherPollSleepMs
        if ($script:pwConnectFailureStreak -gt 0) {
            $backoffMs = [int][Math]::Min(60000, 1000 * [Math]::Pow(2, [Math]::Min($script:pwConnectFailureStreak - 1, 6)))
            $sleepMs = [Math]::Max($sleepMs, $backoffMs)
        }
        $sleepSw = [System.Diagnostics.Stopwatch]::StartNew()
        Start-Sleep -Milliseconds $sleepMs
        $sleepSw.Stop()
        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_TICK_SLEEP' -Message 'Watcher poll tick sleeping.' -Data @{
            tick = $watcherTick
            pollSleepMs = $watcherPollSleepMs
            sleptMs = [int]$sleepSw.ElapsedMilliseconds
            pwSessionOpen = $pwSessionOpen
            pwConnectFailureStreak = $script:pwConnectFailureStreak
            reconnectBackoffMs = if ($script:pwConnectFailureStreak -gt 0) { $sleepMs } else { 0 }
        }
    }
} while ($watcherContinuous)

if ($pwSessionOpen) {
    try {
        Disconnect-PW | Out-Null
        _Watch-WriteJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_DISCONNECT' -Message 'ProjectWise session closed.' -Data @{ continuous = $watcherContinuous; ticks = $watcherTick }
    } catch {
        _Watch-WriteJsonLog -Flush -Level 'Warning' -Code 'WATCH_PW_DISCONNECT_FAILED' -Message 'ProjectWise disconnect failed on watcher exit.' -Data @{ error = [string]$_.Exception.Message }
    }
    $pwSessionOpen = $false
}

exit 0

