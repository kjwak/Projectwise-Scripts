# PW.AuditPoller.psm1
# Responsibility: Query ProjectWise audit trail (dms_audt), resolve folders,
# match against watch roots, and return candidate events for job creation.
# Extracted from the POC Test-AuditEventIngestion.ps1 for production use.

# Dependencies (Core.Results, Core.Runtime, Core.Database) must be imported by the
# caller before this module. Re-importing with -Force here would clobber their
# global-scope exports.

$script:QCRelevantActions = @{
    1002 = 'DOCUMENT_MODIFY'
    1003 = 'DOCUMENT_ATTR'
    1007 = 'DOCUMENT_CIN'
    1012 = 'DOCUMENT_STATE'
    1015 = 'DOCUMENT_VERSION'
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
    foreach ($root in $MatchRoots) {
        if ($FolderPath -like "$root*") { return $true }
    }
    return $false
}

function _AuditPoller-GetSheetsSubpath {
    param([string]$FolderPath, [array]$WatchRootConfigs, [System.Collections.Generic.List[string]]$MatchRoots)
    foreach ($cfg in $WatchRootConfigs) {
        $rootPath = [string]$cfg.path
        $suffix = if ($cfg.sheetsPathFromProject) { [string]$cfg.sheetsPathFromProject } else { 'CADD\Sheets' }
        $testRoots = @($rootPath)
        if ($rootPath -like 'Documents\*') { $testRoots += $rootPath.Substring('Documents\'.Length) }
        else { $testRoots += "Documents\$rootPath" }
        foreach ($tr in $testRoots) {
            if ($FolderPath -like "$tr*" -and $FolderPath -like "*$suffix*") {
                return $true
            }
        }
    }
    return $false
}

function Get-AuditTrailHighWaterMark {
    <#
    .SYNOPSIS
    Reads the most recent successful poll_runs watermark_after from the database.
    Returns a DateTime or $null if no prior run exists.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }
    try {
        $res = Invoke-QCDatabaseScalar -Config $Config -Sql "SELECT TOP 1 watermark_after FROM poll_runs WHERE error_message IS NULL ORDER BY started_at DESC"
        if ($res.IsSuccess -and $res.Data.value) {
            try { return [DateTime]::Parse([string]$res.Data.value) } catch { return $null }
        }
    } catch { }
    return $null
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

    # 1. Query dms_audt
    $sinceStr = $Since.ToString('yyyy-MM-dd HH:mm:ss')
    $sql = "SELECT o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE o_acttime >= '$sinceStr' ORDER BY o_acttime DESC"

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
            watermarkAfter = $sinceStr; durationMs = [int]$sw.ElapsedMilliseconds
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
                    $docToFolder[$dg] = $fp
                    $stats.foldersResolved++
                }
            }
        } catch { }
    }

    # Build parent-GUID to folder map from resolved documents
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
    $watermarkAfter = $sinceStr

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
        if ($actTime -gt $watermarkAfter) { $watermarkAfter = $actTime }

        $candidateType = if ($isWatchMatch) { 'WATCH_MATCH' } else { $null }

        # Write to audit_events table (fire-and-forget)
        if (Test-QCDatabaseEnabled -Config $Config) {
            $objno = 0;  try { $objno = [int]$evt.o_objno } catch { $objno = 0 }
            $userno = 0; try { $userno = [int]$evt.o_userno } catch { $userno = 0 }
            $itemdesc = $null; if (-not ($evt.o_itemdesc -is [DBNull])) { $itemdesc = [string]$evt.o_itemdesc }
            $textparam = $null; if (-not ($evt.o_textparam -is [DBNull])) { $textparam = [string]$evt.o_textparam }

            $insertSql = @"
IF NOT EXISTS (SELECT 1 FROM audit_events WHERE pw_acttime = @acttime AND pw_action = @action AND pw_objguid = @objguid)
INSERT INTO audit_events
    (pw_acttime, pw_action, pw_action_name, pw_objtype, pw_objno, pw_objguid, pw_parentguid, pw_userno, pw_itemname, pw_itemdesc, pw_textparam, resolved_folder, candidate_type)
VALUES
    (@acttime, @action, @actionName, @objtype, @objno, @objguid, @parentguid, @userno, @itemname, @itemdesc, @textparam, @folder, @candidateType)
"@
            $params = @{
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
            try {
                $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $insertSql -Parameters $params
                if ($res.IsSuccess) { $stats.dbWrites++ } else { $stats.dbSkipped++ }
            } catch { $stats.dbSkipped++ }
        }

        if ($isWatchMatch) {
            $candidates += @{
                objGuid        = $objGuid
                parentGuid     = $parentGuid
                actionCode     = $actionCode
                actionName     = $actionName
                itemName       = [string]$evt.o_itemname
                actTime        = $actTime
                resolvedFolder = $resolvedFolder
                isSheetsFolder = $isSheetsFolder
                candidateType  = $candidateType
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

    $counter = 0
    if (Test-Path -LiteralPath $CounterPath) {
        try { $counter = [int](Get-Content -LiteralPath $CounterPath -Raw -ErrorAction Stop).Trim() + 1 }
        catch { $counter = 0 }
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

Export-ModuleMember -Function Invoke-AuditTrailScan, Get-AuditTrailHighWaterMark, Get-AuditPollCycleCounter, Reset-AuditPollCycleCounter
