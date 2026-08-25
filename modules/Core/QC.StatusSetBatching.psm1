Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force -WarningAction SilentlyContinue

function _SSB-ToHashtable([object]$Value) {
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

function _SSB-GetRepoRoot([string]$RepoRoot) {
    if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) { return $RepoRoot }
    return (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
}

function _SSB-ResolveRepoPath([string]$Path, [string]$RepoRoot) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = [string]$Path
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    return (Join-Path (_SSB-GetRepoRoot $RepoRoot) $p)
}

function _SSB-GetQueueRootFromConfig([hashtable]$Config) {
    if (-not $Config) { return '' }
    if ($Config.ContainsKey('queue') -and $Config.queue) {
        $q = _SSB-ToHashtable $Config.queue
        if ($q -and $q.ContainsKey('rootDir') -and $q.rootDir) { return [string]$q.rootDir }
    }
    return ''
}

function _SSB-ResolveDirtyFolderStorePath([string]$Path, [hashtable]$Config, [string]$RepoRoot) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = ([string]$Path).Trim()
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    $queueRoot = _SSB-GetQueueRootFromConfig -Config $Config
    if (-not [string]::IsNullOrWhiteSpace($queueRoot)) {
        return (Join-Path $queueRoot $p.TrimStart('\', '/'))
    }
    return _SSB-ResolveRepoPath -Path $p -RepoRoot $RepoRoot
}

function _SSB-NormalizeFolderKey([string]$FolderPath) {
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return '' }
    if (Get-Command -Name 'Normalize-QCDocumentsFolderPath' -ErrorAction SilentlyContinue) {
        $r = Normalize-QCDocumentsFolderPath -Path $FolderPath
        if ($r.IsSuccess) { return [string]$r.Data.path }
    }
    return ([string]$FolderPath).Trim().TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
}

function _SSB-GetEntryField([object]$Entry, [string]$Name) {
    if (-not $Entry) { return $null }
    if ($Entry -is [hashtable] -and $Entry.ContainsKey($Name)) { return $Entry[$Name] }
    try {
        if ($Entry.PSObject -and $Entry.PSObject.Properties[$Name]) { return $Entry.$Name }
    } catch { }
    return $null
}

function _SSB-ParseUtc([object]$Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    try { return [datetime]::Parse([string]$Value).ToUniversalTime() } catch { return $null }
}

function _SSB-TruncateError([string]$Message) {
    if ([string]::IsNullOrWhiteSpace($Message)) { return '' }
    $text = ([string]$Message).Trim() -replace '\s+', ' '
    if ($text.Length -le 240) { return $text }
    return $text.Substring(0, 240)
}

function _SSB-NewDirtyFolderEntry {
    param(
        [string]$FolderPath,
        [string]$FolderKey,
        [string]$NowUtc,
        [bool]$OneLevelDeep,
        [string]$DatasourceName,
        [string]$TriggerSource,
        [string]$FolderGuid = '',
        [object]$Existing = $null
    )

    $firstSeen = $NowUtc
    $prevFirst = _SSB-GetEntryField $Existing 'firstSeenUtc'
    if (-not $prevFirst) { $prevFirst = _SSB-GetEntryField $Existing 'markedAtUtc' }
    if ($prevFirst) { $firstSeen = [string]$prevFirst }

    $eventCount = 1
    $prevCount = _SSB-GetEntryField $Existing 'eventCount'
    if ($null -ne $prevCount) {
        try { $eventCount = [int]$prevCount + 1 } catch { $eventCount = 1 }
    } elseif ($Existing) { $eventCount = 2 }

    $failureCount = 0
    $prevFailures = _SSB-GetEntryField $Existing 'failureCount'
    if ($null -ne $prevFailures) { try { $failureCount = [int]$prevFailures } catch { } }

    $lastError = _SSB-GetEntryField $Existing 'lastError'
    $lastProcessedUtc = _SSB-GetEntryField $Existing 'lastProcessedUtc'
    $guid = [string]$FolderGuid
    if ([string]::IsNullOrWhiteSpace($guid)) { $guid = [string](_SSB-GetEntryField $Existing 'folderGuid') }

    return @{
        folderPath = $FolderPath
        folderKey = $FolderKey
        folderGuid = if ([string]::IsNullOrWhiteSpace($guid)) { $null } else { $guid }
        firstSeenUtc = $firstSeen
        lastSeenUtc = $NowUtc
        eventCount = $eventCount
        lastProcessedUtc = if ($lastProcessedUtc) { [string]$lastProcessedUtc } else { $null }
        failureCount = $failureCount
        lastError = if ($lastError) { [string]$lastError } else { $null }
        oneLevelDeep = [bool]$OneLevelDeep
        datasourceName = $DatasourceName
        triggerSource = $TriggerSource
    }
}

function Get-QCStatusSetBatchingSettings {
    <#
    .SYNOPSIS
    Resolved statusSetBatching settings from appsettings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$RepoRoot = ''
    )

    $defaults = @{
        enabled = $true
        intervalMinutes = 15
        maxFoldersPerRun = 100
        quietPeriodSeconds = 120
        staleWarningHours = 24
        dirtyFolderStorePath = '_watcher/statusset-dirty-folders.json'
        processOnWatcherStart = $false
    }

    $settings = @{}
    foreach ($key in @($defaults.Keys)) { $settings[$key] = $defaults[$key] }

    $raw = $null
    if ($Config.ContainsKey('statusSetBatching') -and $Config.statusSetBatching) {
        $raw = _SSB-ToHashtable $Config.statusSetBatching
    }
    if ($raw) {
        if ($raw.ContainsKey('enabled') -and $null -ne $raw.enabled) { try { $settings.enabled = [bool]$raw.enabled } catch { } }
        if ($raw.ContainsKey('intervalMinutes') -and $null -ne $raw.intervalMinutes) { try { $settings.intervalMinutes = [int]$raw.intervalMinutes } catch { } }
        if ($raw.ContainsKey('maxFoldersPerRun') -and $null -ne $raw.maxFoldersPerRun) { try { $settings.maxFoldersPerRun = [int]$raw.maxFoldersPerRun } catch { } }
        if ($raw.ContainsKey('quietPeriodSeconds') -and $null -ne $raw.quietPeriodSeconds) { try { $settings.quietPeriodSeconds = [int]$raw.quietPeriodSeconds } catch { } }
        if ($raw.ContainsKey('staleWarningHours') -and $null -ne $raw.staleWarningHours) { try { $settings.staleWarningHours = [int]$raw.staleWarningHours } catch { } }
        if ($raw.ContainsKey('dirtyFolderStorePath') -and $raw.dirtyFolderStorePath) { $settings.dirtyFolderStorePath = [string]$raw.dirtyFolderStorePath }
        if ($raw.ContainsKey('processOnWatcherStart') -and $null -ne $raw.processOnWatcherStart) { try { $settings.processOnWatcherStart = [bool]$raw.processOnWatcherStart } catch { } }
    }

    if ($settings.intervalMinutes -lt 1) { $settings.intervalMinutes = 1 }
    if ($settings.maxFoldersPerRun -lt 1) { $settings.maxFoldersPerRun = 1 }
    if ($settings.quietPeriodSeconds -lt 0) { $settings.quietPeriodSeconds = 0 }
    if ($settings.staleWarningHours -lt 1) { $settings.staleWarningHours = 1 }

    $settings.dirtyFolderStorePath = _SSB-ResolveDirtyFolderStorePath -Path ([string]$settings.dirtyFolderStorePath) -Config $Config -RepoRoot $RepoRoot
    return $settings
}

function Get-QCStatusSetDirtyFolderStorePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$RepoRoot = ''
    )
    $settings = Get-QCStatusSetBatchingSettings -Config $Config -RepoRoot $RepoRoot
    return [string]$settings.dirtyFolderStorePath
}

function Get-QCStatusSetBatchSchedule {
    <#
    .SYNOPSIS
    Next dirty-folder STATUS_SET_GEN batch time from lastBatchRunUtc + intervalMinutes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$RepoRoot = '',
        [datetime]$NowUtc = [datetime]::UtcNow
    )
    $now = $NowUtc
    if ($now.Kind -ne [DateTimeKind]::Utc) { $now = $now.ToUniversalTime() }
    $settings = Get-QCStatusSetBatchingSettings -Config $Config -RepoRoot $RepoRoot
    $dirtyCount = 0
    $lastRun = $null
    try {
        $store = _SSB-ReadDirtyFolderStore -Path ([string]$settings.dirtyFolderStorePath)
        if ($store.folders) { $dirtyCount = @($store.folders.Keys).Count }
        $lastRun = _SSB-ParseUtc $store.lastBatchRunUtc
    } catch { }
    $intervalMin = [int]$settings.intervalMinutes
    $enabled = [bool]$settings.enabled
    $due = $false
    $nextUtc = $null
    if ($enabled) {
        if (-not $lastRun) {
            $due = ($dirtyCount -gt 0)
            if ($due) { $nextUtc = $now }
        } else {
            $nextUtc = $lastRun.AddMinutes($intervalMin)
            if ($now -ge $nextUtc) { $due = $true }
        }
    }
    $minutesUntil = $null
    if ($enabled -and $nextUtc -and -not $due) {
        $minutesUntil = [int][Math]::Max(0, [Math]::Ceiling(($nextUtc - $now).TotalMinutes))
    }
    return New-QCSuccessResult -Code 'STATUSSET_BATCH_SCHEDULE' -Message 'Status-set batch schedule.' -Data @{
        enabled = $enabled
        intervalMinutes = $intervalMin
        quietPeriodSeconds = [int]$settings.quietPeriodSeconds
        dirtyFolderCount = $dirtyCount
        lastBatchRunUtc = $(if ($lastRun) { $lastRun.ToString('o') } else { $null })
        nextBatchUtc = $(if ($nextUtc) { $nextUtc.ToString('o') } else { $null })
        due = $due
        minutesUntil = $minutesUntil
        storePath = [string]$settings.dirtyFolderStorePath
    }
}

function _SSB-NewEmptyStore {
    return @{
        version = 1
        lastBatchRunUtc = $null
        folders = @{}
    }
}

function _SSB-ReadDirtyFolderStore([string]$Path) {
    $empty = _SSB-NewEmptyStore
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $empty }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $empty }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $store = _SSB-NewEmptyStore
        if ($obj.lastBatchRunUtc) { $store.lastBatchRunUtc = [string]$obj.lastBatchRunUtc }
        $folders = @{}
        if ($obj.folders) {
            foreach ($prop in @($obj.folders.PSObject.Properties)) {
                $entry = _SSB-ToHashtable $prop.Value
                if ($entry) { $folders[[string]$prop.Name] = $entry }
            }
        }
        $store.folders = $folders
        return $store
    } catch {
        return $empty
    }
}

function _SSB-WriteDirtyFolderStoreAtomic([string]$Path, [hashtable]$Store) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = ($Path + '.tmp.' + ([guid]::NewGuid().ToString('N')))
    $safeFolders = if ($Store -and $Store.ContainsKey('folders') -and $Store.folders) { $Store.folders } else { @{} }
    $payload = @{
        version = 1
        writtenAtUtc = (Get-QCTimestamp)
        lastBatchRunUtc = if ($Store -and $Store.lastBatchRunUtc) { [string]$Store.lastBatchRunUtc } else { $null }
        folders = $safeFolders
    } | ConvertTo-Json -Depth 12
    $enc = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($tmp, $payload, $enc)
    if (Get-Command -Name '_QCQJ-MoveItemWithRetry' -ErrorAction SilentlyContinue) {
        _QCQJ-MoveItemWithRetry -LiteralPath $tmp -Destination $Path
    } else {
        Move-Item -LiteralPath $tmp -Destination $Path -Force
    }
}

function _SSB-WithDirtyFolderStoreLock {
    param(
        [Parameter(Mandatory)][string]$StorePath,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    $mutexName = 'Global\QCStatusSetDirty_' + ([string][BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($StorePath.ToLowerInvariant())))).Replace('-', '')
    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne(15000)
        if (-not $acquired) {
            return New-QCFailureResult -Code 'STATUSSET_DIRTY_STORE_LOCK_TIMEOUT' -Message 'Timed out waiting for dirty-folder store lock.' -Data @{ storePath = $StorePath }
        }
        $result = & $Action
        if ($result -and $result.PSObject -and ($result.PSObject.Properties['IsSuccess'])) {
            return $result
        }
        return New-QCSuccessResult -Code 'STATUSSET_DIRTY_STORE_OK' -Message 'Store operation completed.' -Data $result
    } finally {
        if ($acquired) {
            try { [void]$mutex.ReleaseMutex() } catch { }
        }
        try { $mutex.Dispose() } catch { }
    }
}

function _SSB-WriteLog {
    param(
        [string]$Code,
        [string]$Message,
        [hashtable]$Data = @{},
        [string]$Level = 'Information',
        [scriptblock]$LogCallback = $null
    )
    if ($LogCallback) {
        try { & $LogCallback $Code $Message $Data $Level | Out-Null } catch { }
        return
    }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $Data
    }
}

function _SSB-RecordDirtyFolderFailure {
    param(
        [string]$StorePath,
        [string]$FolderKey,
        [string]$ErrorMessage
    )
    $now = Get-QCTimestamp
    $shortError = _SSB-TruncateError $ErrorMessage
    [void](_SSB-WithDirtyFolderStoreLock -StorePath $StorePath -Action {
        $store = _SSB-ReadDirtyFolderStore -Path $StorePath
        if (-not $store.folders -or -not $store.folders.ContainsKey($FolderKey)) { return $false }
        $entry = _SSB-ToHashtable $store.folders[$FolderKey]
        if (-not $entry) { return $false }
        $failures = 0
        try { $failures = [int]$entry.failureCount } catch { $failures = 0 }
        $entry.failureCount = $failures + 1
        $entry.lastError = $shortError
        $entry.lastProcessedUtc = $now
        $store.folders[$FolderKey] = $entry
        _SSB-WriteDirtyFolderStoreAtomic -Path $StorePath -Store $store
        return $true
    })
}

function _SSB-RemoveDirtyFolderEntry {
    param(
        [string]$StorePath,
        [string]$FolderKey
    )
    return _SSB-WithDirtyFolderStoreLock -StorePath $StorePath -Action {
        $store = _SSB-ReadDirtyFolderStore -Path $StorePath
        if ($store.folders -and $store.folders.ContainsKey($FolderKey)) {
            $store.folders.Remove($FolderKey) | Out-Null
            _SSB-WriteDirtyFolderStoreAtomic -Path $StorePath -Store $store
        }
        return @($store.folders.Keys).Count
    }
}

function _SSB-TestDirtyFolderStale {
    param(
        [object]$Entry,
        [datetime]$NowUtc,
        [int]$StaleWarningHours
    )
    $firstSeen = _SSB-ParseUtc (_SSB-GetEntryField $Entry 'firstSeenUtc')
    if (-not $firstSeen) { $firstSeen = _SSB-ParseUtc (_SSB-GetEntryField $Entry 'markedAtUtc') }
    if (-not $firstSeen) { return $null }
    $ageHours = ($NowUtc - $firstSeen).TotalHours
    if ($ageHours -lt [double]$StaleWarningHours) { return $null }
    return @{
        ageHours = [math]::Round($ageHours, 2)
        firstSeenUtc = $firstSeen.ToString('o')
    }
}

function _SSB-WriteStaleDirtyFolderWarning {
    param(
        [object]$Entry,
        [string]$FolderKey,
        [hashtable]$StaleInfo,
        [scriptblock]$LogCallback
    )
    $fp = [string](_SSB-GetEntryField $Entry 'folderPath')
    $lastSeen = [string](_SSB-GetEntryField $Entry 'lastSeenUtc')
    if ([string]::IsNullOrWhiteSpace($lastSeen)) { $lastSeen = [string](_SSB-GetEntryField $Entry 'lastEventAtUtc') }
    $eventCount = 0
    try { $eventCount = [int](_SSB-GetEntryField $Entry 'eventCount') } catch { }
    $failureCount = 0
    try { $failureCount = [int](_SSB-GetEntryField $Entry 'failureCount') } catch { }
    $guid = [string](_SSB-GetEntryField $Entry 'folderGuid')

    _SSB-WriteLog -Code 'STATUSSET_DIRTY_FOLDER_STALE' -Message 'Dirty status-set folder exceeded stale warning threshold.' -Level 'Warning' -Data @{
        folderGuid = $guid
        folderPath = $fp
        folderKey = $FolderKey
        firstSeenUtc = [string]$StaleInfo.firstSeenUtc
        lastSeenUtc = $lastSeen
        eventCount = $eventCount
        failureCount = $failureCount
        ageHours = $StaleInfo.ageHours
    } -LogCallback $LogCallback
}

function Mark-StatusSetDirtyFolder {
    <#
    .SYNOPSIS
    Marks a ProjectWise Sheets folder dirty for deferred STATUS_SET_GEN evaluation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [string]$DatasourceName = '',
        [bool]$OneLevelDeep = $true,
        [string]$TriggerSource = 'audit_trail',
        [string]$FolderGuid = '',
        [string]$RepoRoot = '',
        [scriptblock]$LogCallback = $null
    )

    $settings = Get-QCStatusSetBatchingSettings -Config $Config -RepoRoot $RepoRoot
    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'STATUSSET_DIRTY_SKIPPED' -Message 'Status-set batching disabled.' -Data @{ folderPath = $FolderPath; marked = $false }
    }

    $storePath = [string]$settings.dirtyFolderStorePath
    $folderKey = _SSB-NormalizeFolderKey $FolderPath
    if (-not $folderKey) {
        return New-QCFailureResult -Code 'STATUSSET_DIRTY_INVALID_FOLDER' -Message 'Folder path is empty.' -Data @{ folderPath = $FolderPath }
    }

    $now = Get-QCTimestamp
    $lockRes = _SSB-WithDirtyFolderStoreLock -StorePath $storePath -Action {
        $store = _SSB-ReadDirtyFolderStore -Path $storePath
        if (-not $store.folders) { $store.folders = @{} }
        $isNew = -not $store.folders.ContainsKey($folderKey)
        $existing = if ($isNew) { $null } else { _SSB-ToHashtable $store.folders[$folderKey] }

        $entry = _SSB-NewDirtyFolderEntry -FolderPath $FolderPath -FolderKey $folderKey -NowUtc $now `
            -OneLevelDeep:$OneLevelDeep -DatasourceName $DatasourceName -TriggerSource $TriggerSource `
            -FolderGuid $FolderGuid -Existing $existing
        $store.folders[$folderKey] = $entry
        _SSB-WriteDirtyFolderStoreAtomic -Path $storePath -Store $store
        return @{ isNew = $isNew; entry = $entry; dirtyCount = @($store.folders.Keys).Count }
    }
    if (-not $lockRes.IsSuccess) { return $lockRes }

    $d = $lockRes.Data
    _SSB-WriteLog -Code 'STATUSSET_DIRTY_FOLDER_MARKED' -Message 'Sheets folder marked dirty for deferred status-set evaluation.' -Data @{
        folder = $FolderPath
        folderKey = $folderKey
        folderGuid = if ($FolderGuid) { $FolderGuid } else { $null }
        isNew = [bool]$d.isNew
        dirtyFolderCount = [int]$d.dirtyCount
        eventCount = [int]$d.entry.eventCount
        oneLevelDeep = [bool]$OneLevelDeep
        triggerSource = $TriggerSource
    } -LogCallback $LogCallback

    return New-QCSuccessResult -Code 'STATUSSET_DIRTY_FOLDER_MARKED' -Message 'Folder marked dirty.' -Data @{
        folderPath = $FolderPath
        folderKey = $folderKey
        marked = $true
        isNew = [bool]$d.isNew
        dirtyFolderCount = [int]$d.dirtyCount
        entry = $d.entry
    }
}

function Invoke-StatusSetFolderEvaluation {
    <#
    .SYNOPSIS
    Folder-level STATUS_SET_GEN evaluation: PW scan, manifest gate, dedupe, enqueue.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory)]
        [hashtable]$StatusRule,
        [string]$DatasourceName = '',
        [bool]$OneLevelDeep = $true,
        [string]$TriggerSource = 'statusset_batch',
        [string]$ScanReason = 'statusset_batch',
        [bool]$DryRun = $false,
        [scriptblock]$LogCallback = $null
    )

    $fp = [string]$FolderPath
    $result = @{
        folderPath = $fp
        evaluated = $false
        enqueued = $false
        duplicates = $false
        accepted = $false
        skippedReason = $null
        gateReason = $null
        pairedCount = 0
        jobId = $null
    }

    $inFlightRes = Test-QCStatusSetJobInFlight -Config $Config -SourceFolder $fp
    if ($inFlightRes.IsSuccess -and [bool]$inFlightRes.Data.inFlight) {
        $result.skippedReason = 'in_flight'
        $result.jobId = [string]$inFlightRes.Data.jobId
        return New-QCSuccessResult -Code 'STATUSSET_FOLDER_EVALUATED' -Message 'Skipped: STATUS_SET_GEN already in flight.' -Data $result
    }

    if (-not (Get-Command -Name 'Get-StatusSetPWFolderState' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'STATUSSET_FOLDER_EVAL_UNAVAILABLE' -Message 'Get-StatusSetPWFolderState unavailable.' -Data $result
    }

    $pwFolder = $fp
    if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
        $pwFolder = ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp
    }

    $state = Get-StatusSetPWFolderState -FolderPath $pwFolder -OneLevelDeep:$OneLevelDeep
    $result.evaluated = $true
    $result.pairedCount = [int]$state.pairedCount

    if ([int]$state.pairedCount -le 0) {
        $result.skippedReason = if ([int]$state.pdfCount -gt 0 -or [int]$state.dgnCount -gt 0) { 'no_pairs' } else { 'no_docs' }
        return New-QCSuccessResult -Code 'STATUSSET_FOLDER_EVALUATED' -Message 'Folder evaluated; no paired sheets.' -Data $result
    }

    $gateRes = Test-StatusSetWatcherShouldEnqueue -Config $Config -SourceFolder $fp -FolderState $state
    if (-not $gateRes.IsSuccess) {
        return New-QCFailureResult -Code 'STATUSSET_FOLDER_GATE_FAILED' -Message $gateRes.Message -Data $result
    }
    if (-not [bool]$gateRes.Data.shouldEnqueue) {
        $result.skippedReason = 'already_current'
        $result.gateReason = [string]$gateRes.Data.gateReason
        return New-QCSuccessResult -Code 'STATUSSET_FOLDER_EVALUATED' -Message 'Folder evaluated; status set already current.' -Data $result
    }

    $candidate = @{
        path = $fp
        fileName = '_folder_'
        description = ''
        detectedAtUtc = (Get-QCTimestamp)
        sourceFolder = $fp
        datasourceName = $DatasourceName
        groupKey = ('STATUS_SET_GEN|' + $fp).ToLowerInvariant()
        folderStateHash = [string]$state.folderStateHash
        oneLevelDeep = $OneLevelDeep
        triggerSource = $TriggerSource
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

    $jobRes = New-QCJobObject -Candidate $candidate -Rule $StatusRule -Config $Config
    if (-not $jobRes.IsSuccess) {
        return New-QCFailureResult -Code 'STATUSSET_FOLDER_JOB_CREATE_FAILED' -Message $jobRes.Message -Data $result
    }

    $job = [hashtable]$jobRes.Data.job
    $result.accepted = $true
    $result.jobId = [string]$job['id']

    $dupRes = Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $Config
    if (-not $dupRes.IsSuccess) {
        return New-QCFailureResult -Code 'STATUSSET_FOLDER_DEDUPE_FAILED' -Message $dupRes.Message -Data $result
    }

    $wouldDedupe = [bool]$dupRes.Data.isDuplicate
    if ($wouldDedupe) {
        $result.duplicates = $true
        $result.skippedReason = 'duplicate'
        return New-QCSuccessResult -Code 'STATUSSET_FOLDER_EVALUATED' -Message 'Folder evaluated; duplicate job exists.' -Data $result
    }

    if (-not $DryRun) {
        $enqRes = Add-QCQueueJob -Job $job -Config $Config
        if (-not $enqRes.IsSuccess) {
            return New-QCFailureResult -Code 'STATUSSET_FOLDER_ENQUEUE_FAILED' -Message $enqRes.Message -Data $result
        }
        $result.enqueued = $true
    } else {
        $result.enqueued = $false
        $result.skippedReason = 'dry_run'
    }

    return New-QCSuccessResult -Code 'STATUSSET_FOLDER_EVALUATED' -Message 'Folder evaluated; STATUS_SET_GEN enqueued.' -Data $result
}

function Invoke-StatusSetDirtyFolderBatch {
    <#
    .SYNOPSIS
    Processes dirty Sheets folders when the batch interval elapses.
    Bounded work per run (maxFoldersPerRun); folder failures do not stop the batch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$StatusRule,
        [string]$DatasourceName = '',
        [bool]$DryRun = $false,
        [string]$RepoRoot = '',
        [switch]$Force,
        [scriptblock]$LogCallback = $null,
        [scriptblock]$TestFolderEvaluationOverride = $null
    )

    $batchSw = [System.Diagnostics.Stopwatch]::StartNew()
    $batchStats = @{
        foldersConsidered = 0
        foldersProcessed = 0
        foldersSucceeded = 0
        foldersFailed = 0
        foldersSkippedQuietPeriod = 0
        jobsQueued = 0
        elapsedMs = 0
        remaining = 0
        # Legacy aliases retained for callers/tests
        processed = 0
        succeeded = 0
        failed = 0
        enqueued = 0
        duplicates = 0
        skippedQuiet = 0
        cleared = 0
    }

    $settings = Get-QCStatusSetBatchingSettings -Config $Config -RepoRoot $RepoRoot
    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'STATUSSET_BATCH_SKIPPED' -Message 'Status-set batching disabled.' -Data $batchStats
    }

    $storePath = [string]$settings.dirtyFolderStorePath
    $nowUtc = (Get-Date).ToUniversalTime()
    $intervalSec = [int]$settings.intervalMinutes * 60

    $due = [bool]$Force.IsPresent
    if (-not $due) {
        $storePeek = _SSB-ReadDirtyFolderStore -Path $storePath
        $lastRun = _SSB-ParseUtc $storePeek.lastBatchRunUtc
        if (-not $lastRun) { $due = $true }
        elseif (($nowUtc - $lastRun).TotalSeconds -ge [double]$intervalSec) { $due = $true }
    }

    if (-not $due) {
        return New-QCSuccessResult -Code 'STATUSSET_BATCH_NOT_DUE' -Message 'Batch interval not elapsed.' -Data $batchStats
    }

    $runStartedUtc = (Get-QCTimestamp)
    _SSB-WriteLog -Code 'STATUSSET_BATCH_INTERVAL_START' -Message 'Status-set dirty-folder batch run started.' -Data @{
        intervalMinutes = [int]$settings.intervalMinutes
        quietPeriodSeconds = [int]$settings.quietPeriodSeconds
        maxFoldersPerRun = [int]$settings.maxFoldersPerRun
        staleWarningHours = [int]$settings.staleWarningHours
        forced = [bool]$Force.IsPresent
    } -LogCallback $LogCallback

    $planRes = _SSB-WithDirtyFolderStoreLock -StorePath $storePath -Action {
        $store = _SSB-ReadDirtyFolderStore -Path $storePath
        if (-not $store.folders) { $store.folders = @{} }
        $store.lastBatchRunUtc = $runStartedUtc

        $quietSec = [double]$settings.quietPeriodSeconds
        $eligible = [System.Collections.Generic.List[object]]::new()
        $skippedQuiet = 0
        $considered = @($store.folders.Keys).Count
        $staleWarnings = [System.Collections.Generic.List[object]]::new()

        foreach ($key in @($store.folders.Keys)) {
            $entry = _SSB-ToHashtable $store.folders[$key]
            if (-not $entry) { continue }

            $stale = _SSB-TestDirtyFolderStale -Entry $entry -NowUtc $nowUtc -StaleWarningHours ([int]$settings.staleWarningHours)
            if ($stale) {
                [void]$staleWarnings.Add(@{ key = [string]$key; entry = $entry; stale = $stale })
            }

            $lastEvent = _SSB-ParseUtc (_SSB-GetEntryField $entry 'lastSeenUtc')
            if (-not $lastEvent) { $lastEvent = _SSB-ParseUtc (_SSB-GetEntryField $entry 'lastEventAtUtc') }
            if ($lastEvent -and $quietSec -gt 0 -and (($nowUtc - $lastEvent).TotalSeconds -lt $quietSec)) {
                $skippedQuiet++
                continue
            }
            [void]$eligible.Add(@{ key = [string]$key; entry = $entry; lastEvent = $lastEvent })
        }

        $sorted = @($eligible | Sort-Object { if ($_.lastEvent) { $_.lastEvent } else { [datetime]::MinValue } })
        $toProcess = @($sorted | Select-Object -First ([int]$settings.maxFoldersPerRun))

        _SSB-WriteDirtyFolderStoreAtomic -Path $storePath -Store $store
        return @{
            skippedQuiet = $skippedQuiet
            considered = $considered
            eligible = $toProcess
            staleWarnings = @($staleWarnings)
            remaining = @($store.folders.Keys).Count
        }
    }

    if (-not $planRes.IsSuccess) {
        return $planRes
    }

    $batchStats.foldersConsidered = [int]$planRes.Data.considered
    $batchStats.foldersSkippedQuietPeriod = [int]$planRes.Data.skippedQuiet
    $batchStats.skippedQuiet = [int]$planRes.Data.skippedQuiet
    $toProcess = @($planRes.Data.eligible)

    foreach ($warn in @($planRes.Data.staleWarnings)) {
        _SSB-WriteStaleDirtyFolderWarning -Entry $warn.entry -FolderKey ([string]$warn.key) -StaleInfo $warn.stale -LogCallback $LogCallback
    }

    foreach ($item in $toProcess) {
        $key = [string]$item.key
        $entry = _SSB-ToHashtable $item.entry
        if (-not $entry) { continue }
        $fp = [string]$entry.folderPath
        $oneLevelDeep = $true
        try { if ($null -ne $entry.oneLevelDeep) { $oneLevelDeep = [bool]$entry.oneLevelDeep } } catch { }
        $triggerSource = if ($entry.triggerSource) { [string]$entry.triggerSource } else { 'statusset_batch' }
        $ds = if ($entry.datasourceName) { [string]$entry.datasourceName } elseif ($DatasourceName) { $DatasourceName } else { '' }

        $batchStats.foldersProcessed++
        $batchStats.processed++
        _SSB-WriteLog -Code 'STATUSSET_BATCH_FOLDER_START' -Message 'Evaluating dirty folder for STATUS_SET_GEN.' -Data @{
            folder = $fp
            folderKey = $key
            oneLevelDeep = $oneLevelDeep
            triggerSource = $triggerSource
        } -LogCallback $LogCallback

        $evalRes = $null
        try {
            if ($TestFolderEvaluationOverride) {
                $evalRes = & $TestFolderEvaluationOverride $Config $fp $StatusRule $ds $oneLevelDeep $triggerSource $DryRun
            } else {
                $evalRes = Invoke-StatusSetFolderEvaluation -Config $Config -FolderPath $fp -StatusRule $StatusRule `
                    -DatasourceName $ds -OneLevelDeep:$oneLevelDeep -TriggerSource $triggerSource -DryRun:$DryRun -LogCallback $LogCallback
            }
        } catch {
            $batchStats.foldersFailed++
            $batchStats.failed++
            $errMsg = _SSB-TruncateError $_.Exception.Message
            _SSB-RecordDirtyFolderFailure -StorePath $storePath -FolderKey $key -ErrorMessage $errMsg
            _SSB-WriteLog -Code 'STATUSSET_BATCH_FOLDER_FAILED' -Message $errMsg -Level 'Warning' -Data @{
                folder = $fp
                folderKey = $key
                code = 'STATUSSET_FOLDER_EVAL_EXCEPTION'
            } -LogCallback $LogCallback
            continue
        }

        if ($evalRes.IsSuccess) {
            $batchStats.foldersSucceeded++
            $batchStats.succeeded++
            if ($evalRes.Data.enqueued) {
                $batchStats.jobsQueued++
                $batchStats.enqueued++
            }
            if ($evalRes.Data.duplicates) { $batchStats.duplicates++ }

            $clearRes = _SSB-RemoveDirtyFolderEntry -StorePath $storePath -FolderKey $key
            if ($clearRes.IsSuccess) {
                $batchStats.cleared++
                $batchStats.remaining = [int]$clearRes.Data
            }

            _SSB-WriteLog -Code 'STATUSSET_BATCH_FOLDER_EVALUATED' -Message 'Dirty folder evaluated successfully.' -Data @{
                folder = $fp
                folderKey = $key
                skippedReason = $evalRes.Data.skippedReason
                gateReason = $evalRes.Data.gateReason
                enqueued = [bool]$evalRes.Data.enqueued
                duplicates = [bool]$evalRes.Data.duplicates
                pairedCount = [int]$evalRes.Data.pairedCount
                jobId = $evalRes.Data.jobId
            } -LogCallback $LogCallback
        } else {
            $batchStats.foldersFailed++
            $batchStats.failed++
            $errMsg = _SSB-TruncateError $evalRes.Message
            _SSB-RecordDirtyFolderFailure -StorePath $storePath -FolderKey $key -ErrorMessage $errMsg
            _SSB-WriteLog -Code 'STATUSSET_BATCH_FOLDER_FAILED' -Message $errMsg -Level 'Warning' -Data @{
                folder = $fp
                folderKey = $key
                code = [string]$evalRes.Code
            } -LogCallback $LogCallback
        }
    }

    if ($batchStats.remaining -eq 0) {
        $peek = _SSB-ReadDirtyFolderStore -Path $storePath
        $batchStats.remaining = @($peek.folders.Keys).Count
    }

    $batchSw.Stop()
    $batchStats.elapsedMs = [int]$batchSw.ElapsedMilliseconds

    _SSB-WriteLog -Code 'STATUSSET_BATCH_INTERVAL_DONE' -Message 'Status-set dirty-folder batch run completed.' -Data @{
        foldersConsidered = [int]$batchStats.foldersConsidered
        foldersProcessed = [int]$batchStats.foldersProcessed
        foldersSucceeded = [int]$batchStats.foldersSucceeded
        foldersFailed = [int]$batchStats.foldersFailed
        foldersSkippedQuietPeriod = [int]$batchStats.foldersSkippedQuietPeriod
        jobsQueued = [int]$batchStats.jobsQueued
        elapsedMs = [int]$batchStats.elapsedMs
        remaining = [int]$batchStats.remaining
    } -LogCallback $LogCallback

    return New-QCSuccessResult -Code 'STATUSSET_BATCH_INTERVAL_DONE' -Message 'Batch run completed.' -Data $batchStats
}

Export-ModuleMember -Function Get-QCStatusSetBatchingSettings, Get-QCStatusSetDirtyFolderStorePath, Get-QCStatusSetBatchSchedule, Mark-StatusSetDirtyFolder, Invoke-StatusSetFolderEvaluation, Invoke-StatusSetDirtyFolderBatch
