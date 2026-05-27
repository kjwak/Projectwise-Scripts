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
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = ''
    )

    $found = @()
    if ($WatermarkPath -and (Test-Path -LiteralPath $WatermarkPath)) {
        try {
            $raw = (Get-Content -LiteralPath $WatermarkPath -Raw -ErrorAction Stop).Trim()
            $parsed = _AuditPoller-ParseActTime -ActTime $raw
            if ($parsed) { $found += $parsed }
        } catch { }
    }
    $db = Get-AuditTrailHighWaterMarkFromDatabase -Config $Config
    if ($db) { $found += $db }
    if ($found.Count -eq 0) { return $null }
    return ($found | Sort-Object -Descending | Select-Object -First 1)
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
    $since = if ($lastCapture) { $lastCapture } else { $until.AddSeconds(-$LookbackSeconds) }
    $watermarkBefore = if ($lastCapture) { $lastCapture.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }

    return @{
        since           = $since
        until           = $until
        watermarkBefore = $watermarkBefore
        isFirstCapture  = (-not $lastCapture)
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
    }

    # 1. Query dms_audt — exclusive lower bound (last capture), inclusive upper bound (poll start time)
    $sinceStr = $Since.ToString('yyyy-MM-dd HH:mm:ss')
    $untilStr = $Until.ToString('yyyy-MM-dd HH:mm:ss')
    $actionList = (@($script:QCRelevantActions.Keys) | Sort-Object) -join ','
    $sql = "SELECT TOP 500 o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE o_acttime > '$sinceStr' AND o_acttime <= '$untilStr' AND o_objtype = 2 AND o_action IN ($actionList) ORDER BY o_acttime ASC"

    $allEvents = @()
    try {
        $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
        $allEvents = @($result.Rows)
    } catch {
        $sw.Stop()
        return New-QCFailureResult -Code 'AUDIT_QUERY_FAILED' -Message "dms_audt query failed: $($_.Exception.Message)" -Data @{ durationMs = [int]$sw.ElapsedMilliseconds }
    }
    $stats.totalEvents = $allEvents.Count

    # 2. Filter QC-relevant actions
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
    $sinceDt = _AuditPoller-ParseActTime -ActTime $sinceStr

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
        if ($sinceDt -and $actDt -and ($actDt -le $sinceDt)) { continue }

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

    # Batch insert audit_events for this window (best-effort).
    if ($dbRows.Count -gt 0 -and (Test-QCDatabaseEnabled -Config $Config)) {
        $chunkSize = 200
        for ($i = 0; $i -lt $dbRows.Count; $i += $chunkSize) {
            $chunk = @($dbRows[$i..[Math]::Min($i + $chunkSize - 1, $dbRows.Count - 1)])
            $valuesSql = New-Object System.Text.StringBuilder
            $params = @{}
            for ($r = 0; $r -lt $chunk.Count; $r++) {
                $row = $chunk[$r]
                if ($r -gt 0) { [void]$valuesSql.AppendLine(',') }
                [void]$valuesSql.Append(("(@acttime{0},@action{0},@actionName{0},@objtype{0},@objno{0},@objguid{0},@parentguid{0},@userno{0},@itemname{0},@itemdesc{0},@textparam{0},@folder{0},@candidateType{0})" -f $r))
                foreach ($k in @('acttime','action','actionName','objtype','objno','objguid','parentguid','userno','itemname','itemdesc','textparam','folder','candidateType')) {
                    $params[("$k$r")] = $row[$k]
                }
            }

            $insertSql = @"
INSERT INTO audit_events
    (pw_acttime, pw_action, pw_action_name, pw_objtype, pw_objno, pw_objguid, pw_parentguid, pw_userno, pw_itemname, pw_itemdesc, pw_textparam, resolved_folder, candidate_type)
SELECT v.pw_acttime, v.pw_action, v.pw_action_name, v.pw_objtype, v.pw_objno, v.pw_objguid, v.pw_parentguid, v.pw_userno, v.pw_itemname, v.pw_itemdesc, v.pw_textparam, v.resolved_folder, v.candidate_type
FROM (VALUES
$($valuesSql.ToString())
) AS v(pw_acttime, pw_action, pw_action_name, pw_objtype, pw_objno, pw_objguid, pw_parentguid, pw_userno, pw_itemname, pw_itemdesc, pw_textparam, resolved_folder, candidate_type)
WHERE v.pw_objguid IS NOT NULL
  AND NOT EXISTS (
    SELECT 1 FROM audit_events ae
    WHERE ae.pw_acttime = v.pw_acttime AND ae.pw_action = v.pw_action AND ae.pw_objguid = v.pw_objguid
  );
"@
            try {
                $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $insertSql -Parameters $params
                if ($res.IsSuccess) { $stats.dbWrites += [int]$res.Data.rowsAffected } else { $stats.dbSkipped += $chunk.Count }
            } catch {
                $stats.dbSkipped += $chunk.Count
            }
        }
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
