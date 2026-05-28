# PW.AuditPoller.psm1
# Responsibility: Query ProjectWise audit trail (dms_audt), resolve folders,
# match against watch roots, and return candidate events for job creation.
# Extracted from the POC Test-AuditEventIngestion.ps1 for production use.

# Dependencies (Core.Results, Core.Runtime, Core.Database) must be imported by the
# caller before this module. Re-importing with -Force here would clobber their
# global-scope exports.

$script:QCRelevantActions = @{
    1001 = 'DOCUMENT_CREATE'
    1002 = 'DOCUMENT_MODIFY'
    1003 = 'DOCUMENT_ATTR'
    1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN'
    1012 = 'DOCUMENT_STATE'
    1015 = 'DOCUMENT_VERSION'
    1020 = 'DOCUMENT_DELETE'
}

function _AuditPoller-NormalizeFolderPath {
    param([AllowNull()][string]$FolderPath)
    $t = ($FolderPath -as [string]).Trim().TrimEnd('\').Replace('/', '\')
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    if ($t -match '^(?i)documents\\') { return $t }
    return ('Documents\' + $t)
}

function _AuditPoller-BuildMatchRoots {
    param([string[]]$WatchRoots)
    $matchRoots = [System.Collections.Generic.List[string]]::new()
    foreach ($root in $WatchRoots) {
        $matchRoots.Add($root)
        if ($root -like 'Documents\*') {
            $matchRoots.Add($root.Substring('Documents\'.Length))
        } else {
            $matchRoots.Add("Documents\$root")
        }
    }
    return $matchRoots
}

function _AuditPoller-MatchesWatchRoot {
    param([string]$FolderPath, [System.Collections.Generic.List[string]]$MatchRoots)
    $fp = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $fp) { return $false }
    foreach ($root in $MatchRoots) {
        $r = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$root)
        if ($r -and $fp -like "$r*") { return $true }
    }
    return $false
}

function _AuditPoller-GetWatchRootConfigForFolder {
    param([string]$FolderPath, [array]$WatchRootConfigs)
    $fp = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $fp) { return $null }
    foreach ($cfg in @($WatchRootConfigs)) {
        if (-not $cfg) { continue }
        $rootPath = [string]$cfg.path
        if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }
        $testRoots = @($rootPath)
        if ($rootPath -like 'Documents\*') { $testRoots += $rootPath.Substring('Documents\'.Length) }
        else { $testRoots += "Documents\$rootPath" }
        foreach ($tr in $testRoots) {
            $nr = _AuditPoller-NormalizeFolderPath -FolderPath $tr
            if ($nr -and $fp -like "$nr*") { return $cfg }
        }
    }
    return $null
}

function _AuditPoller-GetSheetsSubpath {
    param([string]$FolderPath, [array]$WatchRootConfigs, [System.Collections.Generic.List[string]]$MatchRoots)
    $fp = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $fp) { return $false }
    foreach ($cfg in $WatchRootConfigs) {
        $rootPath = [string]$cfg.path
        $suffix = if ($cfg.sheetsPathFromProject) { [string]$cfg.sheetsPathFromProject } else { 'CADD\Sheets' }
        $testRoots = @($rootPath)
        if ($rootPath -like 'Documents\*') { $testRoots += $rootPath.Substring('Documents\'.Length) }
        else { $testRoots += "Documents\$rootPath" }
        foreach ($tr in $testRoots) {
            $nr = _AuditPoller-NormalizeFolderPath -FolderPath $tr
            if ($nr -and $fp -like "$nr*" -and $fp -like "*$suffix*") {
                return $true
            }
        }
    }
    return $false
}

function _AuditPoller-ParseActTime {
    param([string]$ActTime)
    if ([string]::IsNullOrWhiteSpace($ActTime)) { return $null }
    try { return [DateTime]::Parse($ActTime) } catch { return $null }
}

function Get-AuditTrailHighWaterMarkFromDatabase {
    <#
    .SYNOPSIS
    Reads the most recent successful poll_runs watermark_after from the database.
    Returns a DateTime or $null if no prior run exists.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }
    try {
        $res = Invoke-QCDatabaseScalar -Config $Config -Sql "SELECT TOP 1 watermark_after FROM poll_runs WHERE error_message IS NULL AND watermark_after IS NOT NULL ORDER BY started_at DESC"
        if ($res.IsSuccess -and $res.Data.value) {
            return _AuditPoller-ParseActTime -ActTime ([string]$res.Data.value)
        }
    } catch { }
    return $null
}

function Get-AuditTrailCaptureWatermark {
    <#
    .SYNOPSIS
    Returns the latest audit capture timestamp from the local watermark file and/or poll_runs.

    .DESCRIPTION
    poll_runs.watermark_after is only consulted when audit-capture-watermark.txt exists.
    Deleting queue/_watcher (no file) resets capture to initialLookbackSeconds even if poll_runs
    still has rows — avoids a stale DB watermark after a queue reset.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = ''
    )

    $found = @()
    $watermarkFileExists = ($WatermarkPath -and (Test-Path -LiteralPath $WatermarkPath))
    if ($watermarkFileExists) {
        try {
            $raw = (Get-Content -LiteralPath $WatermarkPath -Raw -ErrorAction Stop).Trim()
            $parsed = _AuditPoller-ParseActTime -ActTime $raw
            if ($parsed) { $found += $parsed }
        } catch { }
        $db = Get-AuditTrailHighWaterMarkFromDatabase -Config $Config
        if ($db) { $found += $db }
    }
    if ($found.Count -eq 0) { return $null }
    return ($found | Sort-Object -Descending | Select-Object -First 1)
}

function _AuditPoller-EscapeSqlLiteral {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace("'", "''")
}

function _AuditPoller-GetAuditPollerInt {
    param([hashtable]$Config, [string]$Key, [int]$Default)
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = $Config.auditPoller
            if ($ap -is [hashtable] -and $ap.ContainsKey($Key) -and $null -ne $ap[$Key]) { return [int]$ap[$Key] }
            if ($ap.PSObject -and $null -ne $ap.$Key) { return [int]$ap.$Key }
        }
    } catch { }
    return $Default
}

function Get-AuditTrailHighWaterMark {
    <#
    .SYNOPSIS
    Back-compat alias for Get-AuditTrailCaptureWatermark (database only when no path given).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = ''
    )
    return Get-AuditTrailCaptureWatermark -Config $Config -WatermarkPath $WatermarkPath
}

function Set-AuditTrailCaptureWatermark {
    <#
    .SYNOPSIS
    Persists the audit capture high-water mark to the local watermark file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WatermarkPath,
        [Parameter(Mandatory)][DateTime]$CapturedThrough
    )

    try {
        $dir = Split-Path -Parent $WatermarkPath
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $value = $CapturedThrough.ToString('yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $WatermarkPath -Value $value -Encoding UTF8 -NoNewline
        return $true
    } catch {
        return $false
    }
}

function Get-AuditTrailPollWindow {
    <#
    .SYNOPSIS
    Computes the audit poll interval: (last capture, now], or lookback on first run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = '',
        [int]$LookbackSeconds = 120
    )

    $until = Get-Date
    $lastCapture = Get-AuditTrailCaptureWatermark -Config $Config -WatermarkPath $WatermarkPath
    $initialLookbackSeconds = $LookbackSeconds
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = $Config.auditPoller
            if ($ap -is [hashtable] -and $ap.ContainsKey('initialLookbackSeconds') -and $null -ne $ap.initialLookbackSeconds) {
                $initialLookbackSeconds = [int]$ap.initialLookbackSeconds
            } elseif ($ap.PSObject -and $ap.initialLookbackSeconds) {
                $initialLookbackSeconds = [int]$ap.initialLookbackSeconds
            }
        }
    } catch { }
    if ($initialLookbackSeconds -lt 1) { $initialLookbackSeconds = $LookbackSeconds }

    $overlapSeconds = $LookbackSeconds
    if ($overlapSeconds -lt 1) { $overlapSeconds = 1 }

    $since = if ($lastCapture) {
        # Steady-state: overlap by lookbackSeconds so late-indexed PW events are not missed
        # (watermark is second-precision; polls are sub-second apart).
        $lastCapture.AddSeconds(-$overlapSeconds)
    } else {
        $until.AddSeconds(-$initialLookbackSeconds)
    }
    if ($since -ge $until) {
        $since = $until.AddSeconds(-$overlapSeconds)
    }
    $watermarkBefore = if ($lastCapture) { $lastCapture.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }

    return @{
        since           = $since
        until           = $until
        watermarkBefore = $watermarkBefore
        isFirstCapture  = (-not $lastCapture)
        lookbackSecondsUsed = if ($lastCapture) { $overlapSeconds } else { $initialLookbackSeconds }
    }
}

function Invoke-AuditTrailScan {
    <#
    .SYNOPSIS
    Queries the PW audit trail for recent QC-relevant events, resolves folders,
    matches against watch roots, and returns structured candidate data.

    .DESCRIPTION
    Production-ready extraction of the POC audit scan logic.
    Returns a QCResult with:
      - Data.events: all QC-relevant audit events
      - Data.candidates: events in watched Sheets folders (ready for job creation)
      - Data.docToFolder: document GUID -> folder path map
      - Data.stats: counts for telemetry
      - Data.watermarkAfter: latest event timestamp (for high-water mark)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][DateTime]$Since,
        [Parameter()][DateTime]$Until = [DateTime]::Now,
        [array]$WatchRootConfigs = @()
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $stats = @{
        totalEvents     = 0
        relevantEvents  = 0
        foldersResolved = 0
        watchMatches    = 0
        sheetsMatches   = 0
        dbWrites        = 0
        dbSkipped       = 0
        pagesFetched    = 0
        eventsTruncated = $false
    }

    # 1. Query dms_audt — paginated ASC so busy servers are not stuck on the oldest TOP 500 only.
    $sinceStr = $Since.ToString('yyyy-MM-dd HH:mm:ss')
    $untilStr = $Until.ToString('yyyy-MM-dd HH:mm:ss')
    $actionList = (@($script:QCRelevantActions.Keys) | Sort-Object) -join ','
    $pageSize = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'pageSize' -Default 500
    $maxPages = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'maxPagesPerPoll' -Default 100
    if ($pageSize -lt 1) { $pageSize = 500 }
    if ($maxPages -lt 1) { $maxPages = 100 }

    $allEvents = [System.Collections.Generic.List[object]]::new()
    $cursorSince = $sinceStr
    $cursorGuid = ''
    $pageNum = 0
    try {
        while ($pageNum -lt $maxPages) {
            $pageNum++
            $lowerBoundSql = if ($cursorGuid) {
                $t = _AuditPoller-EscapeSqlLiteral -Value $cursorSince
                $g = _AuditPoller-EscapeSqlLiteral -Value $cursorGuid
                "(o_acttime > '$t' OR (o_acttime = '$t' AND o_objguid > '$g'))"
            } else {
                "o_acttime > '$(_AuditPoller-EscapeSqlLiteral -Value $cursorSince)'"
            }
            $sql = "SELECT TOP $pageSize o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE $lowerBoundSql AND o_acttime <= '$(_AuditPoller-EscapeSqlLiteral -Value $untilStr)' AND o_objtype = 2 AND o_action IN ($actionList) ORDER BY o_acttime ASC, o_objguid ASC"
            $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
            $batch = @($result.Rows)
            if ($batch.Count -eq 0) { break }
            foreach ($row in $batch) { [void]$allEvents.Add($row) }
            $last = $batch[$batch.Count - 1]
            $cursorSince = [string]$last.o_acttime
            $cursorGuid = [string]$last.o_objguid
            if ($batch.Count -lt $pageSize) { break }
        }
        if ($pageNum -ge $maxPages -and $allEvents.Count -gt 0) {
            $stats.eventsTruncated = $true
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_QUERY_TRUNCATED' -Message "dms_audt page cap reached ($maxPages x $pageSize); re-run will continue from watermark overlap." -Data @{
                    since = $sinceStr; until = $untilStr; pagesFetched = $pageNum; eventsFetched = $allEvents.Count
                }
            }
        }
    } catch {
        $sw.Stop()
        return New-QCFailureResult -Code 'AUDIT_QUERY_FAILED' -Message "dms_audt query failed: $($_.Exception.Message)" -Data @{ durationMs = [int]$sw.ElapsedMilliseconds }
    }
    $stats.pagesFetched = $pageNum
    $stats.totalEvents = $allEvents.Count

    # 2. Filter QC-relevant actions (SQL already filters; keep for safety)
    $relevant = @($allEvents | Where-Object { $script:QCRelevantActions.ContainsKey([int]$_.o_action) })
    $stats.relevantEvents = $relevant.Count

    if ($relevant.Count -eq 0) {
        $sw.Stop()
        return New-QCSuccessResult -Code 'AUDIT_NO_EVENTS' -Message 'No QC-relevant audit events in window.' -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $untilStr; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $untilStr }
        }
    }

    # 3. Resolve folders via document GUIDs (batched)
    $docGuids = @($relevant | ForEach-Object { [string]$_.o_objguid } | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique)
    $docToFolder = @{}
    $folderMap = @{}

    $batchSize = 200
    for ($i = 0; $i -lt $docGuids.Count; $i += $batchSize) {
        $chunk = @($docGuids[$i..[Math]::Min($i + $batchSize - 1, $docGuids.Count - 1)])
        try {
            $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs $chunk -ErrorAction SilentlyContinue)
            foreach ($doc in $docs) {
                $dg = [string]$doc.DocumentGUID
                if (-not $dg) { continue }
                $fp = $null
                if ($doc.FolderPath) { $fp = [string]$doc.FolderPath }
                elseif ($doc.FullPath) {
                    $full = [string]$doc.FullPath
                    $fp = [System.IO.Path]::GetDirectoryName($full) -replace '/', '\'
                }
                if ($fp) {
                    $canonical = _AuditPoller-NormalizeFolderPath -FolderPath $fp
                    if ($canonical) {
                        $docToFolder[$dg] = $canonical
                        $stats.foldersResolved++
                    }
                }
            }
        } catch { }
    }

    # Build parent-GUID to folder map from resolved documents
    $dbRows = @()

    foreach ($evt in $relevant) {
        $og = [string]$evt.o_objguid
        $pg = [string]$evt.o_parentguid
        if ($og -and $docToFolder.ContainsKey($og) -and $pg -and -not $folderMap.ContainsKey($pg)) {
            $folderMap[$pg] = $docToFolder[$og]
        }
    }

    # 4. Match against watch roots
    $watchRoots = @()
    if ($WatchRootConfigs.Count -gt 0) {
        $watchRoots = @($WatchRootConfigs | ForEach-Object { [string]$_.path })
    } elseif ($Config.projectWise -and $Config.projectWise.watchList -and $Config.projectWise.watchList.roots) {
        $WatchRootConfigs = @($Config.projectWise.watchList.roots)
        $watchRoots = @($WatchRootConfigs | ForEach-Object { [string]$_.path })
    }
    $matchRoots = _AuditPoller-BuildMatchRoots -WatchRoots $watchRoots

    $candidates = @()
    $watermarkAfter = $untilStr

    foreach ($evt in $relevant) {
        $actionCode = [int]$evt.o_action
        $actionName = $script:QCRelevantActions[$actionCode]
        $objGuid = [string]$evt.o_objguid
        $parentGuid = [string]$evt.o_parentguid

        $resolvedFolder = $null
        if ($docToFolder.ContainsKey($objGuid)) { $resolvedFolder = $docToFolder[$objGuid] }
        elseif ($folderMap.ContainsKey($parentGuid)) { $resolvedFolder = $folderMap[$parentGuid] }

        $isWatchMatch = $false
        $isSheetsFolder = $false
        if ($resolvedFolder) {
            $isWatchMatch = _AuditPoller-MatchesWatchRoot -FolderPath $resolvedFolder -MatchRoots $matchRoots
            if ($isWatchMatch) {
                $stats.watchMatches++
                $isSheetsFolder = _AuditPoller-GetSheetsSubpath -FolderPath $resolvedFolder -WatchRootConfigs $WatchRootConfigs -MatchRoots $matchRoots
                if ($isSheetsFolder) { $stats.sheetsMatches++ }
            }
        }

        $actTime = [string]$evt.o_acttime
        $actDt = _AuditPoller-ParseActTime -ActTime $actTime
        if ($actDt -and ($actTime -gt $watermarkAfter)) { $watermarkAfter = $actTime }

        $candidateType = if ($isWatchMatch) { 'WATCH_MATCH' } else { $null }

        # Prepare audit_events rows for a set-based insert (fire-and-forget).
        if (Test-QCDatabaseEnabled -Config $Config) {
            $objno = 0;  try { $objno = [int]$evt.o_objno } catch { $objno = 0 }
            $userno = 0; try { $userno = [int]$evt.o_userno } catch { $userno = 0 }
            $itemdesc = $null; if (-not ($evt.o_itemdesc -is [DBNull])) { $itemdesc = [string]$evt.o_itemdesc }
            $textparam = $null; if (-not ($evt.o_textparam -is [DBNull])) { $textparam = [string]$evt.o_textparam }

            $dbRows += @{
                acttime       = $actTime
                action        = $actionCode
                actionName    = $actionName
                objtype       = [int]$evt.o_objtype
                objno         = $objno
                objguid       = $objGuid
                parentguid    = $parentGuid
                userno        = $userno
                itemname      = [string]$evt.o_itemname
                itemdesc      = $itemdesc
                textparam     = $textparam
                folder        = $resolvedFolder
                candidateType = $candidateType
            }
        }

        $enableQcPrepend = $false
        $enableQcCommentSync = $false
        $enableStatusSet = $false
        $watchRootPath = $null
        if ($isWatchMatch) {
            $rootCfg = _AuditPoller-GetWatchRootConfigForFolder -FolderPath $resolvedFolder -WatchRootConfigs $WatchRootConfigs
            if ($rootCfg) {
                try { if ($rootCfg.enableQcPrepend) { $enableQcPrepend = [bool]$rootCfg.enableQcPrepend } } catch { }
                try { if ($rootCfg.enableQcCommentSync) { $enableQcCommentSync = [bool]$rootCfg.enableQcCommentSync } } catch { }
                try { if ($rootCfg.enableStatusSet) { $enableStatusSet = [bool]$rootCfg.enableStatusSet } } catch { }
                if ($rootCfg.path) { $watchRootPath = [string]$rootCfg.path }
            }
            $candidates += @{
                objGuid              = $objGuid
                parentGuid           = $parentGuid
                actionCode           = $actionCode
                actionName           = $actionName
                itemName             = [string]$evt.o_itemname
                actTime              = $actTime
                resolvedFolder       = $resolvedFolder
                isSheetsFolder       = $isSheetsFolder
                candidateType        = $candidateType
                enableQcPrepend      = $enableQcPrepend
                enableQcCommentSync  = $enableQcCommentSync
                enableStatusSet      = $enableStatusSet
                watchRoot            = $watchRootPath
            }
        }
    }

    $userNumbersToSync = [System.Collections.Generic.HashSet[int]]::new()

    # Batch insert audit_events for this window (best-effort).
    if ($dbRows.Count -gt 0) {
        try {
            $dbRes = Write-QCAuditEventRows -Config $Config -Rows $dbRows
            if ($dbRes.IsSuccess -and $dbRes.Data) {
                $stats.dbWrites += [int]$dbRes.Data.written
                $stats.dbSkipped += [int]$dbRes.Data.skipped
            } else {
                $stats.dbSkipped += $dbRows.Count
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_WRITE_FAILED' -Message "Write-QCAuditEventRows failed: $($dbRes.Message)" -Data @{
                        code = [string]$dbRes.Code
                        rowCount = $dbRows.Count
                    }
                }
            }
        } catch {
            $stats.dbSkipped += $dbRows.Count
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_WRITE_EXCEPTION' -Message $_.Exception.Message -Data @{ rowCount = $dbRows.Count }
            }
        }
        foreach ($row in $dbRows) {
            $u = 0
            try { $u = [int]$row.userno } catch { $u = 0 }
            if ($u -gt 0) { [void]$userNumbersToSync.Add($u) }
        }
    }

    if ($userNumbersToSync.Count -gt 0 -and (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) {
        try {
            Sync-PWUserDirectory -Config $Config -UserNumbers @($userNumbersToSync) -MaxUsers 25 | Out-Null
        } catch { }
    }

    $sw.Stop()
    return New-QCSuccessResult -Code 'AUDIT_SCAN_OK' -Message "Audit scan complete: $($stats.relevantEvents) relevant, $($stats.watchMatches) matched." -Data @{
        events         = $relevant
        candidates     = $candidates
        docToFolder    = $docToFolder
        stats          = $stats
        watermarkAfter = $watermarkAfter
        durationMs     = [int]$sw.ElapsedMilliseconds
        pollWindow     = @{ since = $sinceStr; until = $untilStr }
    }
}

function Get-AuditPollCycleCounter {
    <#
    .SYNOPSIS
    Reads/increments a cycle counter from a file to determine when to run reconciliation.
    Returns the current cycle number (1-based).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CounterPath)

    # 1-based: first tick returns 1 so cycle 0 does not force an immediate full-folder reconciliation.
    $counter = 1
    if (Test-Path -LiteralPath $CounterPath) {
        try { $counter = [int](Get-Content -LiteralPath $CounterPath -Raw -ErrorAction Stop).Trim() + 1 }
        catch { $counter = 1 }
    }
    try {
        $dir = Split-Path -Parent $CounterPath
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Set-Content -LiteralPath $CounterPath -Value ([string]$counter) -Encoding UTF8
    } catch { }
    return $counter
}

function Reset-AuditPollCycleCounter {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CounterPath)
    try { Set-Content -LiteralPath $CounterPath -Value '0' -Encoding UTF8 } catch { }
}

Export-ModuleMember -Function Invoke-AuditTrailScan, Get-AuditTrailHighWaterMark, Get-AuditTrailHighWaterMarkFromDatabase, Get-AuditTrailCaptureWatermark, Set-AuditTrailCaptureWatermark, Get-AuditTrailPollWindow, Get-AuditPollCycleCounter, Reset-AuditPollCycleCounter
