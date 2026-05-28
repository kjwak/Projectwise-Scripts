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

# Full PW action map for audit_events ingestion (telemetry is unfiltered; triggers use QCRelevantActions only).
$script:AuditActionNames = @{
    1    = 'FOLDER_CREATE';     2    = 'FOLDER_MODIFY';     3    = 'FOLDER_WFLOW'
    4    = 'FOLDER_DELETE';     5    = 'FOLDER_STATE'
    1001 = 'DOCUMENT_CREATE';   1002 = 'DOCUMENT_MODIFY';   1003 = 'DOCUMENT_ATTR'
    1004 = 'DOCUMENT_FILE_ADD'; 1005 = 'DOCUMENT_FILE_REM'; 1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN';     1008 = 'DOCUMENT_VIEW';     1009 = 'DOCUMENT_CHOUT'
    1010 = 'DOCUMENT_CPOUT';   1011 = 'DOCUMENT_GOUT';     1012 = 'DOCUMENT_STATE'
    1013 = 'DOCUMENT_FINAL_S'; 1014 = 'DOCUMENT_FINAL_R'; 1015 = 'DOCUMENT_VERSION'
    1016 = 'DOCUMENT_MOVE';    1020 = 'DOCUMENT_DELETE';    1022 = 'DOCUMENT_FREE'
    1027 = 'DOCUMENT_IMPORT'; 3001 = 'USER_LOGIN';        3002 = 'USER_LOGOUT'
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

function _AuditPoller-TryAdvanceWatermarkAfter {
    param(
        [AllowNull()][string]$Current,
        [AllowNull()][string]$Candidate
    )
    if ([string]::IsNullOrWhiteSpace($Candidate)) {
        if ([string]::IsNullOrWhiteSpace($Current)) { return $null }
        return $Current
    }
    if ([string]::IsNullOrWhiteSpace($Current)) { return $Candidate }
    $curDt = _AuditPoller-ParseActTime -ActTime $Current
    $candDt = _AuditPoller-ParseActTime -ActTime $Candidate
    if ($candDt -and (-not $curDt -or $candDt -gt $curDt)) { return $Candidate }
    return $Current
}

function _AuditPoller-GetRowValue {
    param($Row, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Row) { return $null }
    try {
        if ($Row -is [System.Data.DataRow]) {
            foreach ($col in $Row.Table.Columns) {
                if ([string]::Equals([string]$col.ColumnName, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                    $v = $Row[$col]
                    if ($v -is [DBNull]) { return $null }
                    return $v
                }
            }
        }
        foreach ($prop in $Row.PSObject.Properties) {
            if ([string]::Equals([string]$prop.Name, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                $v = $prop.Value
                if ($v -is [DBNull]) { return $null }
                return $v
            }
        }
    } catch { }
    return $null
}

function _AuditPoller-FormatActTime {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [DBNull]) { return $null }
    if ($Value -is [DateTime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    $s = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $parsed = _AuditPoller-ParseActTime -ActTime $s
    if ($parsed) { return $parsed.ToString('yyyy-MM-dd HH:mm:ss') }
    return $s
}

function _AuditPoller-GetSqlResultRows {
    param($Result)
    if ($null -eq $Result) { return @() }
    try {
        if ($Result.PSObject.Properties.Name -contains 'Rows' -and $null -ne $Result.Rows) {
            return @($Result.Rows)
        }
    } catch { }
    if ($Result -is [System.Data.DataTable]) { return @($Result.Rows) }
    return @($Result)
}

function _AuditPoller-GetActionCode {
    param($Row)
    $v = _AuditPoller-GetRowValue -Row $Row -Name 'o_action'
    if ($null -eq $v) { return 0 }
    try { return [int]$v } catch { return 0 }
}

function _AuditPoller-GetActionName {
    param([int]$ActionCode)
    if ($script:QCRelevantActions.ContainsKey($ActionCode)) { return $script:QCRelevantActions[$ActionCode] }
    if ($script:AuditActionNames.ContainsKey($ActionCode)) { return $script:AuditActionNames[$ActionCode] }
    return "UNKNOWN_$ActionCode"
}

function _AuditPoller-NewAuditEventDbRow {
    param(
        $Evt,
        [AllowNull()][string]$ResolvedFolder = $null,
        [AllowNull()][string]$CandidateType = $null
    )
    $actionCode = _AuditPoller-GetActionCode -Row $Evt
    $objno = 0;  try { $objno = [int](_AuditPoller-GetRowValue -Row $Evt -Name 'o_objno') } catch { $objno = 0 }
    $userno = 0; try { $userno = [int](_AuditPoller-GetRowValue -Row $Evt -Name 'o_userno') } catch { $userno = 0 }
    $itemdesc = _AuditPoller-GetRowValue -Row $Evt -Name 'o_itemdesc'
    $textparam = _AuditPoller-GetRowValue -Row $Evt -Name 'o_textparam'
    $objtype = 0; try { $objtype = [int](_AuditPoller-GetRowValue -Row $Evt -Name 'o_objtype') } catch { $objtype = 0 }
    return @{
        acttime       = (_AuditPoller-FormatActTime -Value (_AuditPoller-GetRowValue -Row $Evt -Name 'o_acttime'))
        action        = $actionCode
        actionName    = (_AuditPoller-GetActionName -ActionCode $actionCode)
        objtype       = $objtype
        objno         = $objno
        objguid       = [string](_AuditPoller-GetRowValue -Row $Evt -Name 'o_objguid')
        parentguid    = [string](_AuditPoller-GetRowValue -Row $Evt -Name 'o_parentguid')
        userno        = $userno
        itemname      = [string](_AuditPoller-GetRowValue -Row $Evt -Name 'o_itemname')
        itemdesc      = if ($null -eq $itemdesc) { $null } else { [string]$itemdesc }
        textparam     = if ($null -eq $textparam) { $null } else { [string]$textparam }
        folder        = $ResolvedFolder
        candidateType = $CandidateType
    }
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
        dbRowsPrepared  = 0
        dbRowsNullGuid  = 0
        pagesFetched    = 0
        eventsTruncated = $false
        dbLastError     = $null
    }

    # 1. Query dms_audt — paginated ASC so busy servers are not stuck on the oldest TOP 500 only.
    # SQL bounds use the poll window (watcher local clock). PW o_acttime must be comparable to these strings
    # (typically: run the watcher in the same timezone as the PW/SQL datasource).
    $sinceStr = $Since.ToString('yyyy-MM-dd HH:mm:ss')
    $queryUntilStr = $Until.ToString('yyyy-MM-dd HH:mm:ss')
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
            $untilStr = $queryUntilStr
            $lowerBoundSql = if ($cursorGuid) {
                $t = _AuditPoller-EscapeSqlLiteral -Value $cursorSince
                $g = _AuditPoller-EscapeSqlLiteral -Value $cursorGuid
                "(o_acttime > '$t' OR (o_acttime = '$t' AND o_objguid > '$g'))"
            } else {
                "o_acttime > '$(_AuditPoller-EscapeSqlLiteral -Value $cursorSince)'"
            }
            $sql = "SELECT TOP $pageSize o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE $lowerBoundSql AND o_acttime <= '$(_AuditPoller-EscapeSqlLiteral -Value $untilStr)' ORDER BY o_acttime ASC, o_objguid ASC"
            $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
            $batch = @(_AuditPoller-GetSqlResultRows -Result $result)
            if ($batch.Count -eq 0) { break }
            foreach ($row in $batch) { [void]$allEvents.Add($row) }
            $last = $batch[$batch.Count - 1]
            $cursorSince = _AuditPoller-FormatActTime -Value (_AuditPoller-GetRowValue -Row $last -Name 'o_acttime')
            if ([string]::IsNullOrWhiteSpace($cursorSince)) { break }
            $cursorGuid = [string](_AuditPoller-GetRowValue -Row $last -Name 'o_objguid')
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

    if ($allEvents.Count -eq 0) {
        $sw.Stop()
        return New-QCSuccessResult -Code 'AUDIT_NO_EVENTS' -Message 'No audit events in window.' -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $queryUntilStr; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
        }
    }

    # 2. Ingest every fetched row into audit_events (no QC/watch/action filtering).
    $maxPwActTime = $null
    $watermarkAfter = $queryUntilStr
    $dbRows = @()
    if (Test-QCDatabaseEnabled -Config $Config) {
        foreach ($evt in $allEvents) {
            $actTime = _AuditPoller-FormatActTime -Value (_AuditPoller-GetRowValue -Row $evt -Name 'o_acttime')
            if ($actTime) { $maxPwActTime = _AuditPoller-TryAdvanceWatermarkAfter -Current $maxPwActTime -Candidate $actTime }
            $dbRows += (_AuditPoller-NewAuditEventDbRow -Evt $evt)
        }
    } else {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_INGEST_DISABLED' -Message 'database.enabled is false; audit_events not written.' -Data @{ eventsFetched = $allEvents.Count }
        }
    }

    $stats.dbRowsPrepared = $dbRows.Count
    $stats.dbRowsNullGuid = @($dbRows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.objguid) }).Count

    $userNumbersToSync = [System.Collections.Generic.HashSet[int]]::new()
    if ($dbRows.Count -gt 0) {
        try {
            $dbRes = Write-QCAuditEventRows -Config $Config -Rows $dbRows
            if ($dbRes.IsSuccess -and $dbRes.Data) {
                $stats.dbWrites += [int]$dbRes.Data.written
                $stats.dbSkipped += [int]$dbRes.Data.skipped
                if ($dbRes.Data.lastError) { $stats.dbLastError = [string]$dbRes.Data.lastError }
            } else {
                $stats.dbSkipped += $dbRows.Count
                $stats.dbLastError = [string]$dbRes.Message
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_WRITE_FAILED' -Message "Write-QCAuditEventRows failed: $($dbRes.Message)" -Data @{
                        code = [string]$dbRes.Code
                        rowCount = $dbRows.Count
                        nullGuid = $stats.dbRowsNullGuid
                    }
                }
            }
        } catch {
            $stats.dbSkipped += $dbRows.Count
            $stats.dbLastError = [string]$_.Exception.Message
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_WRITE_EXCEPTION' -Message $_.Exception.Message -Data @{ rowCount = $dbRows.Count; nullGuid = $stats.dbRowsNullGuid }
            }
        }
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            $ingestLevel = if ($stats.dbWrites -gt 0) { 'Information' } elseif ($stats.dbRowsPrepared -gt 0) { 'Warning' } else { 'Information' }
            Write-QCJsonLog -Flush -Level $ingestLevel -Code 'AUDIT_EVENTS_INGEST' -Message "audit_events ingest: $($stats.dbWrites) written, $($stats.dbSkipped) skipped/duplicate." -Data @{
                eventsFetched = $stats.totalEvents
                rowsPrepared  = $stats.dbRowsPrepared
                rowsNullGuid  = $stats.dbRowsNullGuid
                written       = $stats.dbWrites
                skipped       = $stats.dbSkipped
                lastError     = $stats.dbLastError
            }
        }
        foreach ($row in $dbRows) {
            $u = 0
            try { $u = [int]$row.userno } catch { $u = 0 }
            if ($u -gt 0) { [void]$userNumbersToSync.Add($u) }
        }
    }

    if ($maxPwActTime) { $stats.maxPwActTime = $maxPwActTime }

    # 3. QC trigger pipeline — filter applies only to job candidates, not audit_events ingestion.
    $relevant = @($allEvents | Where-Object { $script:QCRelevantActions.ContainsKey((_AuditPoller-GetActionCode -Row $_)) })
    $stats.relevantEvents = $relevant.Count

    if ($relevant.Count -eq 0) {
        if ($userNumbersToSync.Count -gt 0 -and (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) {
            try { Sync-PWUserDirectory -Config $Config -UserNumbers @($userNumbersToSync) -MaxUsers 25 | Out-Null } catch { }
        }
        $sw.Stop()
        return New-QCSuccessResult -Code 'AUDIT_SCAN_OK' -Message "Audit ingest complete: $($stats.totalEvents) fetched, 0 QC-relevant for triggers." -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $watermarkAfter; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
        }
    }

    # 4. Resolve folders via document GUIDs (batched) — QC-relevant document events only
    $docGuids = @($relevant | ForEach-Object { [string](_AuditPoller-GetRowValue -Row $_ -Name 'o_objguid') } | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique)
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
    foreach ($evt in $relevant) {
        $og = [string](_AuditPoller-GetRowValue -Row $evt -Name 'o_objguid')
        $pg = [string](_AuditPoller-GetRowValue -Row $evt -Name 'o_parentguid')
        if ($og -and $docToFolder.ContainsKey($og) -and $pg -and -not $folderMap.ContainsKey($pg)) {
            $folderMap[$pg] = $docToFolder[$og]
        }
    }

    # 5. Match against watch roots (job candidates only)
    $watchRoots = @()
    if ($WatchRootConfigs.Count -gt 0) {
        $watchRoots = @($WatchRootConfigs | ForEach-Object { [string]$_.path })
    } elseif ($Config.projectWise -and $Config.projectWise.watchList -and $Config.projectWise.watchList.roots) {
        $WatchRootConfigs = @($Config.projectWise.watchList.roots)
        $watchRoots = @($WatchRootConfigs | ForEach-Object { [string]$_.path })
    }
    $matchRoots = _AuditPoller-BuildMatchRoots -WatchRoots $watchRoots

    $candidates = @()

    foreach ($evt in $relevant) {
        $actionCode = _AuditPoller-GetActionCode -Row $evt
        $actionName = $script:QCRelevantActions[$actionCode]
        $objGuid = [string](_AuditPoller-GetRowValue -Row $evt -Name 'o_objguid')
        $parentGuid = [string](_AuditPoller-GetRowValue -Row $evt -Name 'o_parentguid')

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

        $actTime = [string](_AuditPoller-GetRowValue -Row $evt -Name 'o_acttime')
        $candidateType = if ($isWatchMatch) { 'WATCH_MATCH' } else { $null }

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
                itemName             = [string](_AuditPoller-GetRowValue -Row $evt -Name 'o_itemname')
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
        pollWindow     = @{ since = $sinceStr; until = $queryUntilStr }
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
