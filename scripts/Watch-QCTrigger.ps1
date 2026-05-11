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

function _New-WatchTimingMap {
    return @{
        candidateDiscoveryMs = 0
        metadataCollectionMs = 0
        triggerMatchingMs = 0
        hashCalculationMs = 0
        jobFactoryMs = 0
        dedupeLookupMs = 0
        enqueueWriteMs = 0
        projectWiseScanMs = 0
    }
}

function _Add-WatchTiming {
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Timing,
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [System.Diagnostics.Stopwatch]$Stopwatch
    )
    try {
        if (-not $Timing.Contains($Name)) { $Timing[$Name] = 0 }
        $Timing[$Name] = [int64]$Timing[$Name] + [int64]$Stopwatch.ElapsedMilliseconds
    } catch { }
}

function _Invoke-WatchTimed {
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Timing,
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        return & $ScriptBlock
    } finally {
        $sw.Stop()
        _Add-WatchTiming -Timing $Timing -Name $Name -Stopwatch $sw
    }
}


function _QCW-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCW-ToStableJsonValue([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [System.ValueType]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $ordered = [ordered]@{}
        foreach ($k in @($Value.Keys | Sort-Object { [string]$_ })) {
            $ordered[[string]$k] = _QCW-ToStableJsonValue -Value $Value[$k]
        }
        return $ordered
    }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        foreach ($i in $Value) { $items += (_QCW-ToStableJsonValue -Value $i) }
        return $items
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $ordered = [ordered]@{}
        foreach ($p in @($Value.PSObject.Properties | Sort-Object Name)) {
            $ordered[$p.Name] = _QCW-ToStableJsonValue -Value $p.Value
        }
        return $ordered
    }
    return [string]$Value
}

function _QCW-GetLocalCachePath([hashtable]$Config) {
    $root = $null
    try {
        if ($Config.ContainsKey('queue') -and $Config.queue -and $Config.queue.ContainsKey('rootDir') -and $Config.queue.rootDir) {
            $root = [string]$Config.queue.rootDir
        }
    } catch { }
    if (_QCW-IsNullOrWhiteSpace $root) {
        $root = Join-Path $env:TEMP 'QCQueue'
    }
    return (Join-Path (Join-Path $root '_watcher') 'local-file-cache.json')
}

function _QCW-GetLocalCacheConfigHash([hashtable]$Config) {
    $fingerprint = @{
        schema = 'watcherLocalFileCacheV1'
        filters = if ($Config.ContainsKey('filters')) { $Config.filters } else { $null }
        triggers = if ($Config.ContainsKey('triggers')) { $Config.triggers } else { $null }
    }
    $json = (_QCW-ToStableJsonValue -Value $fingerprint) | ConvertTo-Json -Depth 80 -Compress
    return Get-Sha256TextHex -Text $json
}

function _QCW-ReadLocalFileCache([hashtable]$Config) {
    $path = _QCW-GetLocalCachePath -Config $Config
    $configHash = _QCW-GetLocalCacheConfigHash -Config $Config
    $cache = @{
        schemaVersion = 1
        configHash = $configHash
        path = $path
        files = @{}
        dirty = $false
        loaded = $false
        ignoredReason = ''
    }
    if (-not (Test-Path -LiteralPath $path)) { return $cache }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $cache }
        $obj = ConvertTo-HashtableDeep -Value ($raw | ConvertFrom-Json -ErrorAction Stop)
        if (-not ($obj -is [hashtable])) { return $cache }
        if ([string]$obj.configHash -ne $configHash) {
            $cache.ignoredReason = 'config_hash_changed'
            return $cache
        }
        if ($obj.ContainsKey('files') -and $obj.files -is [hashtable]) {
            $cache.files = $obj.files
        }
        $cache.loaded = $true
    } catch {
        $cache.ignoredReason = 'read_failed'
    }
    return $cache
}

function _QCW-WriteLocalFileCache([hashtable]$Cache) {
    try {
        if (-not $Cache -or -not [bool]$Cache.dirty) { return }
        $path = [string]$Cache.path
        if (_QCW-IsNullOrWhiteSpace $path) { return }
        $dir = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $payload = @{
            schemaVersion = 1
            configHash = [string]$Cache.configHash
            updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
            files = $Cache.files
        }
        $tmp = $path + '.tmp.' + ([guid]::NewGuid().ToString('N'))
        $json = $payload | ConvertTo-Json -Depth 80 -Compress
        $enc = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($tmp, $json, $enc)
        Move-Item -LiteralPath $tmp -Destination $path -Force
    } catch { }
}

function _QCW-GetFileSignature([string]$Path, [int64]$Length, [datetime]$LastWriteTimeUtc) {
    $ticks = 0
    try { $ticks = ([datetime]$LastWriteTimeUtc).ToUniversalTime().Ticks } catch { $ticks = 0 }
    return (([string]$Path).ToLowerInvariant() + '|mtimeTicks=' + [string]$ticks + '|len=' + [string]$Length)
}

function _QCW-GetCacheKey([string]$NormalizedPath) {
    return ([string]$NormalizedPath).ToLowerInvariant()
}

function _QCW-GetLocalCacheEntry([hashtable]$Cache, [string]$Key) {
    if (-not $Cache -or -not ($Cache.files -is [hashtable])) { return $null }
    if ($Cache.files.ContainsKey($Key) -and $Cache.files[$Key] -is [hashtable]) { return $Cache.files[$Key] }
    return $null
}

function _QCW-SetLocalCacheEntry([hashtable]$Cache, [string]$Key, [hashtable]$Entry) {
    if (-not $Cache) { return }
    if (-not ($Cache.files -is [hashtable])) { $Cache.files = @{} }
    $Cache.files[$Key] = $Entry
    $Cache.dirty = $true
}

function _QCW-NewCacheEntry([string]$Signature) {
    return @{
        signature = $Signature
        sha256 = ''
        disposition = ''
        updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
    }
}

function _QCW-SetCacheDisposition([hashtable]$Cache, [string]$Key, [string]$Signature, [string]$Disposition, [string]$Sha256 = '') {
    $entry = _QCW-GetLocalCacheEntry -Cache $Cache -Key $Key
    if (-not $entry -or [string]$entry.signature -ne $Signature) { $entry = _QCW-NewCacheEntry -Signature $Signature }
    if (-not (_QCW-IsNullOrWhiteSpace $Sha256)) { $entry.sha256 = $Sha256 }
    $entry.disposition = $Disposition
    $entry.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
    _QCW-SetLocalCacheEntry -Cache $Cache -Key $Key -Entry $entry
}

function _QCW-SetCacheHash([hashtable]$Cache, [string]$Key, [string]$Signature, [string]$Sha256) {
    if (_QCW-IsNullOrWhiteSpace $Sha256) { return }
    $entry = _QCW-GetLocalCacheEntry -Cache $Cache -Key $Key
    if (-not $entry -or [string]$entry.signature -ne $Signature) { $entry = _QCW-NewCacheEntry -Signature $Signature }
    $entry.sha256 = $Sha256
    $entry.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
    _QCW-SetLocalCacheEntry -Cache $Cache -Key $Key -Entry $entry
}

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

$scriptDir = _Get-ThisScriptDir
$repoRoot = Split-Path -Parent $scriptDir
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = (Join-Path $repoRoot 'appsettings.json')
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Hashing.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Filters.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Triggers.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.JobFactory.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force
$pwConnPath = (Join-Path $repoRoot 'modules\PW.Connection.psm1')
if (-not (Test-Path -LiteralPath $pwConnPath)) {
    throw "PW.Connection.psm1 not found at expected path: $pwConnPath"
}
Import-Module $pwConnPath -Force | Out-Null
Import-Module (Join-Path $repoRoot 'modules\QC.StatusSet.psm1') -Force

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
$watchRunSw = [System.Diagnostics.Stopwatch]::StartNew()
$watchTiming = _New-WatchTimingMap
$watchCounts = @{
    pwFoldersPrepared = 0
    pwFoldersScanned = 0
    pwDocQueries = 0
    pwDocsScanned = 0
    pwTaggedDocs = 0
    localFoldersScanned = 0
    localFilesDiscovered = 0
    hashesCalculated = 0
    hashCacheHits = 0
    hashCacheMisses = 0
    localCacheHits = 0
    localCacheMisses = 0
    localCacheSkips = 0
    localCacheEntries = 0
    triggerRuleCacheUses = 0
    dedupeChecks = 0
    enqueueWrites = 0
}

$orderedRulesRes = Get-OrderedTriggerRules -Config $config
if (-not $orderedRulesRes.IsSuccess) { throw $orderedRulesRes.Message }
$orderedTriggerRules = @($orderedRulesRes.Data.rules)

$localFileCache = _QCW-ReadLocalFileCache -Config $config
try { $watchCounts.localCacheEntries = [int]@($localFileCache.files.Keys).Count } catch { }

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

# ProjectWise watchList processing (STATUS_SET_GEN and/or QC_PREPEND).
# This must run even when STATUS_SET_GEN rules are disabled, because QC_PREPEND can be PW-triggered too.
if ($statusSetRules.Count -ge 0) {
    # ProjectWise sources (watchList) — read-only.
    if ($hasPwWatchList) {
        $pwScanSw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $pwCfg = ConvertTo-HashtableDeep -Value $config.projectWise
            $ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
            $credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
            $watchList = ConvertTo-HashtableDeep -Value $pwCfg.watchList

            # Re-import here to avoid any odd module/session state where exports are not visible.
            Import-Module $pwConnPath -Force | Out-Null
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
                    $enableStatusSet = $false
                    if ($rh.ContainsKey('enableStatusSet')) { try { $enableStatusSet = [bool]$rh.enableStatusSet } catch { $enableStatusSet = $false } }
                    $discovered = @(Find-PWSheetsFoldersUnderRoot -RootPath $rootPath -SheetsSuffix $suffix -DatasourceName $ds -ProjectDepth $projectDepth)
                    foreach ($d in $discovered) {
                        $d['EnableQcPrepend'] = $enableQcPrepend
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
                    $enableStatusSet = $false
                    if ($fh.ContainsKey('enableStatusSet')) { try { $enableStatusSet = [bool]$fh.enableStatusSet } catch { $enableStatusSet = $false } }
                    $full = ($root.TrimEnd('\') + '\' + $path.TrimStart('\')).Trim()
                    $pwFolders += @(@{
                        DatasourceName = $ds
                        FolderPath = $full
                        OneLevelDeep = $oneLevelDeep
                        EnableQcPrepend = $enableQcPrepend
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
            $watchCounts.pwFoldersPrepared = [int]$pwFolders.Count

            # Select-Object -First 1 already yields a single hashtable (or $null). Avoid @() which forces object[].
            $statusRuleObj = ($statusSetRules | Sort-Object -Property priority | Select-Object -First 1)
            foreach ($entry in @($pwFolders)) {
                try {
                    $fp = [string]$entry.FolderPath
                    if ([string]::IsNullOrWhiteSpace($fp)) { continue }
                    $watchCounts.pwFoldersScanned = [int]$watchCounts.pwFoldersScanned + 1

                    $oneLevelDeep = $false
                    $enableQcPrepend = $false
                    $enableStatusSet = $false
                    try { $oneLevelDeep = [bool]$entry.OneLevelDeep } catch { $oneLevelDeep = $false }
                    try { $enableQcPrepend = [bool]$entry.EnableQcPrepend } catch { $enableQcPrepend = $false }
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
                        $state = _Invoke-WatchTimed -Timing $watchTiming -Name 'metadataCollectionMs' -ScriptBlock { Get-StatusSetPWFolderState -FolderPath (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp) -OneLevelDeep:$oneLevelDeep }
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_STATUSSET_SCAN_DONE' -Message 'PW status-set folder query completed.' -Data @{
                            folder = $fp
                            oneLevelDeep = $oneLevelDeep
                            pdfCount = [int]$state.pdfCount
                            dgnCount = [int]$state.dgnCount
                            pairedCount = [int]$state.pairedCount
                        }
                        if ([int]$state.pairedCount -gt 0) {
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
                                    detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
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
                                        lastWriteTimeUtc = ([DateTime]::UtcNow.ToString('o'))
                                    }
                                }

                                $jobRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'jobFactoryMs' -ScriptBlock { New-QCJobObject -Candidate $candidate -Rule $statusRuleObj -Config $config }
                                if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
                                $job = [hashtable]$jobRes.Data.job

                                $accepted++
                                $watchCounts.dedupeChecks = [int]$watchCounts.dedupeChecks + 1
                                $dupRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'dedupeLookupMs' -ScriptBlock { Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config }
                                if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                                $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                                $wouldEnqueue = (-not $wouldDedupe)
                                $enqueueSkippedReason = $null
                                if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
                                elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

                                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW folder change candidate accepted (STATUS_SET_GEN).' -Data @{
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
                                    $watchCounts.enqueueWrites = [int]$watchCounts.enqueueWrites + 1
                                    $enqRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'enqueueWriteMs' -ScriptBlock { Add-QCQueueJob -Job $job -Config $config }
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
                        $watchCounts.pwDocQueries = [int]$watchCounts.pwDocQueries + 1
                        $docs = _Invoke-WatchTimed -Timing $watchTiming -Name 'candidateDiscoveryMs' -ScriptBlock { Get-PWDocumentsInFolder -FolderPath (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp) }
                        $watchCounts.pwDocsScanned = [int]$watchCounts.pwDocsScanned + [int](@($docs).Count)
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
                            $desc = Get-PWDocDescription -Doc $doc
                            if ([string]::IsNullOrWhiteSpace($desc)) { continue }
                            if ($desc.IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }

                            $watchCounts.pwTaggedDocs = [int]$watchCounts.pwTaggedDocs + 1

                            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_TAGGED' -Message 'PW doc has QC_Archivist tag.' -Data @{
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
                                detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
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

                            $watchCounts.triggerRuleCacheUses = [int]$watchCounts.triggerRuleCacheUses + 1
                            $matchRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'triggerMatchingMs' -ScriptBlock { Test-QCTriggerCandidate -Candidate $candidate -Config $config -OrderedRules $orderedTriggerRules }
                            if (-not $matchRes.IsSuccess) { throw $matchRes.Message }
                            if (-not [bool]$matchRes.Data.matched) {
                                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_PW_NO_MATCH' -Message 'PW doc had QC_Archivist but did not match any trigger rule.' -Data @{
                                    path = [string]$candidate.path
                                    fileName = [string]$candidate.fileName
                                    ruleReason = if ($matchRes.Data.ContainsKey('reason')) { [string]$matchRes.Data.reason } else { '' }
                                    candidateDescription = $desc
                                }
                                continue
                            }
                            $ruleObj = $matchRes.Data.rule
                            if ([string]$ruleObj.jobType -ne 'QC_PREPEND') { continue }

                            $jobRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'jobFactoryMs' -ScriptBlock { New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config }
                            if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
                            $job = [hashtable]$jobRes.Data.job

                            $accepted++
                            $watchCounts.dedupeChecks = [int]$watchCounts.dedupeChecks + 1
                            $dupRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'dedupeLookupMs' -ScriptBlock { Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config }
                            if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
                            $wouldDedupe = [bool]$dupRes.Data.isDuplicate
                            $wouldEnqueue = (-not $wouldDedupe)
                            $enqueueSkippedReason = $null
                            if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
                            elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

                            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'PW doc accepted (QC_PREPEND via description tag).' -Data @{
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
                                $watchCounts.enqueueWrites = [int]$watchCounts.enqueueWrites + 1
                                $enqRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'enqueueWriteMs' -ScriptBlock { Add-QCQueueJob -Job $job -Config $config }
                                if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
                                $enqueued++
                            } elseif ($wouldDedupe) { $duplicates++ }
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

            Disconnect-PW | Out-Null
        } catch {
            $errors++
            Write-QCJsonLog -Flush -Level 'Error' -Code 'WATCH_PW_ERROR' -Message 'ProjectWise watchList processing failed.' -Data @{ errorMessage = [string]$_.Exception.Message; scriptStackTrace = [string]$_.ScriptStackTrace }
        } finally {
            $pwScanSw.Stop()
            _Add-WatchTiming -Timing $watchTiming -Name 'projectWiseScanMs' -Stopwatch $pwScanSw
        }
    }

    foreach ($folder in $watchFolders) {
        try {
            if (-not (Test-Path -LiteralPath $folder)) { continue }
            $watchCounts.localFoldersScanned = [int]$watchCounts.localFoldersScanned + 1
            $state = _Invoke-WatchTimed -Timing $watchTiming -Name 'metadataCollectionMs' -ScriptBlock { Get-StatusSetLocalFolderState -RootFolder $folder }
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
                detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
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
                    lastWriteTimeUtc = ([DateTime]::UtcNow.ToString('o'))
                }
            }

            $jobRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'jobFactoryMs' -ScriptBlock { New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config }
            if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
            $job = [hashtable]$jobRes.Data.job

            $accepted++
            $watchCounts.dedupeChecks = [int]$watchCounts.dedupeChecks + 1
            $dupRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'dedupeLookupMs' -ScriptBlock { Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config }
            if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
            $wouldDedupe = [bool]$dupRes.Data.isDuplicate
            $wouldEnqueue = (-not $wouldDedupe)
            $enqueueSkippedReason = $null
            if ($wouldDedupe) { $enqueueSkippedReason = 'duplicate' }
            elseif ($isDryRun) { $enqueueSkippedReason = 'dryRun' }

            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Folder change candidate accepted (STATUS_SET_GEN).' -Data @{
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
                $watchCounts.enqueueWrites = [int]$watchCounts.enqueueWrites + 1
                $enqRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'enqueueWriteMs' -ScriptBlock { Add-QCQueueJob -Job $job -Config $config }
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
}

$fileItems = _Invoke-WatchTimed -Timing $watchTiming -Name 'candidateDiscoveryMs' -ScriptBlock {
    $items = @()
    foreach ($folder in $watchFolders) {
        if (-not (Test-Path -LiteralPath $folder)) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_FOLDER_MISSING' -Message 'Watch folder missing.' -Data @{ folder = $folder }
            continue
        }
        $watchCounts.localFoldersScanned = [int]$watchCounts.localFoldersScanned + 1
        $items += Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue
    }
    return @($items)
}

if ($MaxFiles -gt 0) { $fileItems = @($fileItems | Select-Object -First $MaxFiles) }
$watchCounts.localFilesDiscovered = [int]$fileItems.Count

foreach ($fi in $fileItems) {
    try {
        $metadataSw = [System.Diagnostics.Stopwatch]::StartNew()
        $pathRes = Normalize-QCPath -Path ([string]$fi.FullName)
        if (-not $pathRes.IsSuccess) { throw $pathRes.Message }
        $normPath = [string]$pathRes.Data.path
        $cacheKey = _QCW-GetCacheKey -NormalizedPath $normPath
        $fileSignature = _QCW-GetFileSignature -Path $normPath -Length ([int64]$fi.Length) -LastWriteTimeUtc $fi.LastWriteTimeUtc
        $cacheEntry = _QCW-GetLocalCacheEntry -Cache $localFileCache -Key $cacheKey
        $cacheHit = ($cacheEntry -and [string]$cacheEntry.signature -eq $fileSignature)
        if ($cacheHit) {
            $watchCounts.localCacheHits = [int]$watchCounts.localCacheHits + 1
            $cachedDisposition = [string]$cacheEntry.disposition
            if (-not $isDryRun -and @('filtered','ignored','enqueued','duplicate') -contains $cachedDisposition) {
                $watchCounts.localCacheSkips = [int]$watchCounts.localCacheSkips + 1
                $metadataSw.Stop()
                _Add-WatchTiming -Timing $watchTiming -Name 'metadataCollectionMs' -Stopwatch $metadataSw
                continue
            }
        } else {
            $watchCounts.localCacheMisses = [int]$watchCounts.localCacheMisses + 1
            $cacheEntry = $null
        }

        $allowRes = Test-QCPathAllowed -CandidatePath $normPath -Config $config
        if (-not $allowRes.IsSuccess) { throw $allowRes.Message }
        if (-not [bool]$allowRes.Data.allowed) {
            $filtered++
            _QCW-SetCacheDisposition -Cache $localFileCache -Key $cacheKey -Signature $fileSignature -Disposition 'filtered'
            $metadataSw.Stop()
            _Add-WatchTiming -Timing $watchTiming -Name 'metadataCollectionMs' -Stopwatch $metadataSw
            continue
        }

        $candidate = @{
            path = $normPath
            fileName = [string]$fi.Name
            description = '' # local filesystem has no PW description; triggers should use filename/path/extension.
            detectedAtUtc = ([DateTime]::UtcNow.ToString('o'))
            file = @{
                fullName = [string]$fi.FullName
                length = [int64]$fi.Length
                lastWriteTimeUtc = $fi.LastWriteTimeUtc.ToString('o')
            }
        }
        $sfRes = Normalize-QCPath -Path ([string]$fi.DirectoryName)
        if (-not $sfRes.IsSuccess) { throw $sfRes.Message }
        $candidate.sourceFolder = [string]$sfRes.Data.path
        $metadataSw.Stop()
        _Add-WatchTiming -Timing $watchTiming -Name 'metadataCollectionMs' -Stopwatch $metadataSw

        $watchCounts.triggerRuleCacheUses = [int]$watchCounts.triggerRuleCacheUses + 1
        $matchRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'triggerMatchingMs' -ScriptBlock { Test-QCTriggerCandidate -Candidate $candidate -Config $config -OrderedRules $orderedTriggerRules }
        if (-not $matchRes.IsSuccess) { throw $matchRes.Message }
        if (-not [bool]$matchRes.Data.matched) {
            $ignored++
            if (($ignored % $ignoreSampleEvery) -eq 0) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_IGNORED_SAMPLE' -Message 'Ignored file (no trigger match).' -Data @{
                    path = $normPath
                    fileName = $candidate.fileName
                    ignoredCount = $ignored
                }
            }
            _QCW-SetCacheDisposition -Cache $localFileCache -Key $cacheKey -Signature $fileSignature -Disposition 'ignored'
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
            if ($cacheHit -and $cacheEntry -and -not (_QCW-IsNullOrWhiteSpace $cacheEntry.sha256)) {
                $watchCounts.hashCacheHits = [int]$watchCounts.hashCacheHits + 1
                $candidate.file.sha256 = [string]$cacheEntry.sha256
            } else {
                $watchCounts.hashCacheMisses = [int]$watchCounts.hashCacheMisses + 1
                $watchCounts.hashesCalculated = [int]$watchCounts.hashesCalculated + 1
                $candidate.file.sha256 = _Invoke-WatchTimed -Timing $watchTiming -Name 'hashCalculationMs' -ScriptBlock { Get-Sha256FileHex -Path ([string]$fi.FullName) }
                _QCW-SetCacheHash -Cache $localFileCache -Key $cacheKey -Signature $fileSignature -Sha256 ([string]$candidate.file.sha256)
            }
        } else {
            # grouped folder workflow: establish groupKey = jobType + sourceFolder
            $candidate.groupKey = ($jobType + '|' + [string]$candidate.sourceFolder).ToLowerInvariant()
        }

        $jobRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'jobFactoryMs' -ScriptBlock { New-QCJobObject -Candidate $candidate -Rule $ruleObj -Config $config }
        if (-not $jobRes.IsSuccess) { throw $jobRes.Message }
        $job = [hashtable]$jobRes.Data.job

        $accepted++
        $watchCounts.dedupeChecks = [int]$watchCounts.dedupeChecks + 1
        $dupRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'dedupeLookupMs' -ScriptBlock { Test-QCDuplicateJob -DedupeKey ([string]$job['dedupeKey']) -Config $config }
        if (-not $dupRes.IsSuccess) { throw $dupRes.Message }
        $wouldDedupe = [bool]$dupRes.Data.isDuplicate
        $wouldEnqueue = (-not $wouldDedupe)
        $enqueueSkippedReason = $null
        if ($wouldDedupe) {
            $enqueueSkippedReason = 'duplicate'
        } elseif ($isDryRun) {
            $enqueueSkippedReason = 'dryRun'
        }

        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_ACCEPTED' -Message 'Trigger matched; job accepted.' -Data @{
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
            _QCW-SetCacheDisposition -Cache $localFileCache -Key $cacheKey -Signature $fileSignature -Disposition 'duplicate' -Sha256 ([string]$candidate.file.sha256)
            continue
        }

        $watchCounts.enqueueWrites = [int]$watchCounts.enqueueWrites + 1
        $enqRes = _Invoke-WatchTimed -Timing $watchTiming -Name 'enqueueWriteMs' -ScriptBlock { Add-QCQueueJob -Job $job -Config $config }
        if (-not $enqRes.IsSuccess) { throw $enqRes.Message }
        $enqueued++
        _QCW-SetCacheDisposition -Cache $localFileCache -Key $cacheKey -Signature $fileSignature -Disposition 'enqueued' -Sha256 ([string]$candidate.file.sha256)
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

_QCW-WriteLocalFileCache -Cache $localFileCache
try { $watchCounts.localCacheEntries = [int]@($localFileCache.files.Keys).Count } catch { }
$watchRunSw.Stop()

Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_DONE' -Message 'Watch run completed.' -Data @{
    dryRun = $isDryRun
    scanned = $fileItems.Count
    filtered = $filtered
    ignored = $ignored
    matched = $matched
    accepted = $accepted
    duplicates = $duplicates
    skippedStatusSetCurrent = $skippedStatusSetCurrent
    enqueued = $enqueued
    errors = $errors
    elapsedMs = [int64]$watchRunSw.ElapsedMilliseconds
    phaseMs = $watchTiming
    phaseCounts = $watchCounts
}

exit 0

