<#
.SYNOPSIS
Proof-of-concept: poll PW audit trail, write events to SQL Server, compare with queue.

.DESCRIPTION
End-to-end validation of the audit-trail-driven architecture:
  1. Connect to ProjectWise
  2. Query dms_audt for recent events (watermark approach)
  3. Filter for QC-relevant action types (check-in, attribute, state change)
  4. Resolve to watch folders
  5. Write matching events to the audit_events table in SQL Server
  6. Compare with current queue contents
  7. Report findings

This does NOT modify ProjectWise. It only reads audit data and writes to QC_Pipeline DB.

.PARAMETER Hours
Lookback window in hours. Default 4.

.PARAMETER DryRun
Show what would be written without actually writing to the database.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [int]$Hours = 4,
    [switch]$DryRun
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

$modulesDir = Join-Path $repoRoot 'modules'
foreach ($mod in @('Core.Results.psm1', 'Core.Runtime.psm1', 'Core.Database.psm1', 'PW.Connection.psm1')) {
    $modPath = Join-Path $modulesDir $mod
    if (-not (Test-Path -LiteralPath $modPath)) {
        Write-Host "ERROR: Module not found: $modPath" -ForegroundColor Red
        return
    }
    try { Import-Module $modPath -Force -ErrorAction Stop }
    catch {
        Write-Host "ERROR: Failed to import $mod -- $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

$QCRelevantActions = @{
    1001 = 'DOCUMENT_CREATE'
    1002 = 'DOCUMENT_MODIFY'
    1003 = 'DOCUMENT_ATTR'
    1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN'
    1012 = 'DOCUMENT_STATE'
    1015 = 'DOCUMENT_VERSION'
    1020 = 'DOCUMENT_DELETE'
}

$AllActionMap = @{
    1    = 'FOLDER_CREATE';     2    = 'FOLDER_MODIFY';     3    = 'FOLDER_WFLOW'
    4    = 'FOLDER_DELETE';     5    = 'FOLDER_STATE'
    1001 = 'DOCUMENT_CREATE';   1002 = 'DOCUMENT_MODIFY';   1003 = 'DOCUMENT_ATTR'
    1004 = 'DOCUMENT_FILE_ADD'; 1005 = 'DOCUMENT_FILE_REM'; 1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN';     1008 = 'DOCUMENT_VIEW';     1009 = 'DOCUMENT_CHOUT'
    1010 = 'DOCUMENT_CPOUT';   1011 = 'DOCUMENT_GOUT';     1012 = 'DOCUMENT_STATE'
    1013 = 'DOCUMENT_FINAL_S'; 1014 = 'DOCUMENT_FINAL_R';  1015 = 'DOCUMENT_VERSION'
    1016 = 'DOCUMENT_MOVE';    1020 = 'DOCUMENT_DELETE';    1022 = 'DOCUMENT_FREE'
    1027 = 'DOCUMENT_IMPORT';  3001 = 'USER_LOGIN';        3002 = 'USER_LOGOUT'
}

# -- Load config ---------------------------------------------------------------

Write-Host "Loading config from: $AppSettingsPath" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $AppSettingsPath)) { throw "appsettings.json not found: $AppSettingsPath" }

function _DeepHashtable ($obj) {
    if ($obj -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = _DeepHashtable $p.Value }
        return $h
    }
    if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
        return @($obj | ForEach-Object { _DeepHashtable $_ })
    }
    return $obj
}

$raw = Get-Content -LiteralPath $AppSettingsPath -Raw -ErrorAction Stop
$config = _DeepHashtable ($raw | ConvertFrom-Json -ErrorAction Stop)

$dbEnabled = Test-QCDatabaseEnabled -Config $config
Write-Host "  database.enabled = $dbEnabled" -ForegroundColor $(if ($dbEnabled) { 'Green' } else { 'DarkYellow' })
if ($DryRun) { Write-Host "  DRY RUN - will not write to database" -ForegroundColor Yellow }

# -- 1. Connect to ProjectWise ------------------------------------------------

Write-Host "`n[1] Connecting to ProjectWise..." -ForegroundColor Yellow
$ds = $null
try {
    $pw = $config.projectWise
    if (-not $pw) { throw "projectWise section not configured" }
    $ds = $null
    if ($pw.datasourceName) { $ds = [string]$pw.datasourceName }
    elseif ($pw.datasource) { $ds = [string]$pw.datasource }
    if (-not $ds) { throw "projectWise.datasourceName not configured" }
    $credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
    $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
    if (-not $credRes.IsSuccess) { throw "Credential error: $($credRes.Message)" }
    $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
    if (-not $connRes.IsSuccess) { throw "Connection error: $($connRes.Message)" }
    Write-Host "  Connected to datasource: $ds" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# -- 2. Query dms_audt with watermark -----------------------------------------

$since = (Get-Date).AddHours(-$Hours).ToString('yyyy-MM-dd HH:mm:ss')
Write-Host "`n[2] Querying dms_audt for events since $since ($Hours hours)..." -ForegroundColor Yellow

$sql = @"
SELECT o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid,
       o_userno, o_itemname, o_itemdesc, o_textparam
FROM dms_audt
WHERE o_acttime >= '$since'
ORDER BY o_acttime DESC
"@

$allEvents = @()
try {
    $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
    $allEvents = @($result.Rows)
    Write-Host "  Total events: $($allEvents.Count)" -ForegroundColor Green
} catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PW | Out-Null
    return
}

# -- 3. Filter for QC-relevant actions ----------------------------------------

Write-Host "`n[3] Filtering for QC-relevant actions..." -ForegroundColor Yellow
$relevant = @($allEvents | Where-Object { $QCRelevantActions.ContainsKey([int]$_.o_action) })
Write-Host "  QC-relevant events: $($relevant.Count) / $($allEvents.Count) total" -ForegroundColor Green

$byAction = @{}
foreach ($evt in $relevant) {
    $name = $QCRelevantActions[[int]$evt.o_action]
    if (-not $byAction.ContainsKey($name)) { $byAction[$name] = 0 }
    $byAction[$name]++
}
foreach ($k in ($byAction.Keys | Sort-Object)) {
    Write-Host "    $k : $($byAction[$k])" -ForegroundColor Cyan
}

# -- 4. Show unique documents affected ----------------------------------------

Write-Host "`n[4] Unique documents affected..." -ForegroundColor Yellow
$docGuids = @($relevant | ForEach-Object { [string]$_.o_objguid } | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique)
Write-Host "  Unique documents: $($docGuids.Count)" -ForegroundColor Green

$sample = @($relevant | Select-Object -First 15)
if ($sample.Count -gt 0) {
    Write-Host "`n  Recent QC-relevant events:" -ForegroundColor Cyan
    Write-Host ("  {0,-20} {1,-18} {2,-50}" -f 'Time', 'Action', 'Document')
    Write-Host ("  {0,-20} {1,-18} {2,-50}" -f '----', '------', '--------')
    foreach ($evt in $sample) {
        $actionName = $QCRelevantActions[[int]$evt.o_action]
        $name = [string]$evt.o_itemname
        if ($name.Length -gt 50) { $name = $name.Substring(0, 47) + '...' }
        $time = [string]$evt.o_acttime
        Write-Host ("  {0,-20} {1,-18} {2,-50}" -f $time, $actionName, $name)
    }
}

# -- 5. Resolve parent folders (sample) ---------------------------------------

Write-Host "`n[5] Resolving folders via document GUIDs..." -ForegroundColor Yellow
$docGuidList = @($relevant | ForEach-Object { [string]$_.o_objguid } | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique)
$folderMap = @{}
$docToFolder = @{}
if ($docGuidList.Count -gt 0) {
    $sampleDocGuids = @($docGuidList | Select-Object -First 50)
    $resolved = 0
    foreach ($dg in $sampleDocGuids) {
        try {
            $doc = Get-PWDocumentsByGUIDs -DocumentGUIDs @($dg) -ErrorAction SilentlyContinue
            if ($doc) {
                $fp = $null
                if ($doc.FolderPath) { $fp = [string]$doc.FolderPath }
                elseif ($doc.FullPath) {
                    $full = [string]$doc.FullPath
                    $fp = [System.IO.Path]::GetDirectoryName($full) -replace '/', '\'
                }
                if ($fp) {
                    $docToFolder[$dg] = $fp
                    # Also map the parent GUID for the event-level folder resolution
                    $matchingEvt = $relevant | Where-Object { [string]$_.o_objguid -eq $dg } | Select-Object -First 1
                    if ($matchingEvt) {
                        $pg = [string]$matchingEvt.o_parentguid
                        if ($pg -and -not $folderMap.ContainsKey($pg)) { $folderMap[$pg] = $fp }
                    }
                    $resolved++
                }
            }
        } catch { }
    }
    Write-Host "  Resolved $resolved / $($sampleDocGuids.Count) document folders" -ForegroundColor Green
    $uniqueFolders = @($docToFolder.Values | Select-Object -Unique)
    foreach ($f in ($uniqueFolders | Select-Object -First 10)) {
        Write-Host "    $f" -ForegroundColor DarkCyan
    }
}

# -- 6. Match against watch folder config -------------------------------------

Write-Host "`n[6] Matching against configured watch folders..." -ForegroundColor Yellow
$watchRoots = @()
try {
    if ($config.projectWise -and $config.projectWise.watchList -and $config.projectWise.watchList.roots) {
        $watchRoots = @($config.projectWise.watchList.roots | ForEach-Object { [string]$_.path })
    }
} catch { }

$matchCount = 0
$matchedFolders = @{}
$allResolvedPaths = @($folderMap.Values) + @($docToFolder.Values) | Select-Object -Unique

# Build alternate roots: strip leading "Documents\" if present, and add it if not,
# so matching works regardless of whether the API returns the prefix.
$matchRoots = [System.Collections.Generic.List[string]]::new()
foreach ($root in $watchRoots) {
    $matchRoots.Add($root)
    if ($root -like 'Documents\*') {
        $matchRoots.Add($root.Substring('Documents\'.Length))
    } else {
        $matchRoots.Add("Documents\$root")
    }
}

foreach ($path in $allResolvedPaths) {
    foreach ($root in $matchRoots) {
        if ($path -like "$root*") {
            $matchCount++
            if (-not $matchedFolders.ContainsKey($path)) { $matchedFolders[$path] = 0 }
            $matchedFolders[$path]++
            break
        }
    }
}
Write-Host "  Watch roots configured: $($watchRoots.Count)" -ForegroundColor Gray
Write-Host "  Folders matching watch roots: $matchCount / $($folderMap.Count)" -ForegroundColor $(if ($matchCount -gt 0) { 'Green' } else { 'DarkYellow' })
if ($matchedFolders.Count -gt 0) {
    foreach ($f in ($matchedFolders.Keys | Select-Object -First 5)) {
        Write-Host "    [MATCH] $f" -ForegroundColor Green
    }
}

# -- 7. Write to audit_events table -------------------------------------------

Write-Host "`n[7] Writing to audit_events table..." -ForegroundColor Yellow
$written = 0
$skipped = 0

if (-not $dbEnabled) {
    Write-Host "  SKIPPED: database not enabled" -ForegroundColor DarkYellow
    $skipped = $relevant.Count
} elseif ($DryRun) {
    Write-Host "  SKIPPED (dry run): would write $($relevant.Count) events" -ForegroundColor Yellow
    $skipped = $relevant.Count
} else {
    foreach ($evt in $relevant) {
        $actionCode = [int]$evt.o_action
        $actionName = $QCRelevantActions[$actionCode]
        $objGuid = [string]$evt.o_objguid
        $parentGuid = [string]$evt.o_parentguid
        $resolvedFolder = $null
        if ($docToFolder.ContainsKey($objGuid)) { $resolvedFolder = $docToFolder[$objGuid] }
        elseif ($folderMap.ContainsKey($parentGuid)) { $resolvedFolder = $folderMap[$parentGuid] }

        $candidateType = $null
        if ($resolvedFolder) {
            foreach ($root in $matchRoots) {
                if ($resolvedFolder -like "$root*") { $candidateType = 'WATCH_MATCH'; break }
            }
        }

        $sql = @"
IF NOT EXISTS (SELECT 1 FROM audit_events WHERE pw_acttime = @acttime AND pw_action = @action AND pw_objguid = @objguid)
INSERT INTO audit_events
    (pw_acttime, pw_action, pw_action_name, pw_objtype, pw_objno, pw_objguid, pw_parentguid, pw_userno, pw_itemname, pw_itemdesc, pw_textparam, resolved_folder, candidate_type)
VALUES
    (@acttime, @action, @actionName, @objtype, @objno, @objguid, @parentguid, @userno, @itemname, @itemdesc, @textparam, @folder, @candidateType)
"@
        $objno = 0;  try { $objno = [int]$evt.o_objno } catch { $objno = 0 }
        $userno = 0; try { $userno = [int]$evt.o_userno } catch { $userno = 0 }
        $itemdesc = $null; if (-not ($evt.o_itemdesc -is [DBNull])) { $itemdesc = [string]$evt.o_itemdesc }
        $textparam = $null; if (-not ($evt.o_textparam -is [DBNull])) { $textparam = [string]$evt.o_textparam }

        $params = @{
            acttime       = [string]$evt.o_acttime
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
            $res = Invoke-QCDatabaseNonQuery -Config $config -Sql $sql -Parameters $params
            if ($res.IsSuccess) { $written++ } else { $skipped++ }
        } catch { $skipped++ }
    }
    Write-Host "  Written: $written   Skipped (dupes/errors): $skipped" -ForegroundColor Green
}

# -- 8. Compare with queue ----------------------------------------------------

Write-Host "`n[8] Comparing with current queue..." -ForegroundColor Yellow
try {
    Import-Module (Join-Path $modulesDir 'Queue\QC.Queue.Json.psm1') -Force -ErrorAction Stop
    $stats = Get-QCQueueStats -Config $config
    if ($stats.IsSuccess) {
        $st = $stats.Data.states
        Write-Host "  Queue: pending=$([int]$st.pending) running=$([int]$st.running) succeeded=$([int]$st.succeeded) failed=$([int]$st.failed)" -ForegroundColor Cyan
    }
} catch {
    Write-Host "  Queue check skipped: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# -- 9. Check audit_events in DB ----------------------------------------------

Write-Host "`n[9] Audit events now in database..." -ForegroundColor Yellow
if ($dbEnabled -and -not $DryRun) {
    $countRes = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT COUNT(*) FROM audit_events"
    if ($countRes.IsSuccess) {
        Write-Host "  Total audit_events rows: $($countRes.Data.value)" -ForegroundColor Green
    }
    $typeRes = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT pw_action_name, COUNT(*) AS cnt, COUNT(DISTINCT pw_objguid) AS docs FROM audit_events GROUP BY pw_action_name ORDER BY cnt DESC"
    if ($typeRes.IsSuccess -and $typeRes.Data.rowCount -gt 0) {
        Write-Host "`n  Events by type:" -ForegroundColor Cyan
        foreach ($row in $typeRes.Data.table.Rows) {
            Write-Host ("    {0,-20} {1,5} events  {2,4} documents" -f [string]$row.pw_action_name, [int]$row.cnt, [int]$row.docs)
        }
    }
    $matchRes = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT COUNT(*) FROM audit_events WHERE candidate_type = 'WATCH_MATCH'"
    if ($matchRes.IsSuccess) {
        Write-Host "`n  Events matching watch folders: $($matchRes.Data.value)" -ForegroundColor $(if ([int]$matchRes.Data.value -gt 0) { 'Green' } else { 'DarkYellow' })
    }
} else {
    Write-Host "  Skipped (DB disabled or dry run)" -ForegroundColor DarkYellow
}

# -- Disconnect ----------------------------------------------------------------

Disconnect-PW | Out-Null

# -- Summary -------------------------------------------------------------------

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Lookback:              $Hours hours (since $since)"
Write-Host "Total audit events:    $($allEvents.Count)"
Write-Host "QC-relevant events:    $($relevant.Count)"
Write-Host "Unique documents:      $($docGuids.Count)"
Write-Host "Folders resolved:      $($folderMap.Count)"
Write-Host "Watch folder matches:  $matchCount"
Write-Host "DB writes:             $written"
Write-Host "DB skipped:            $skipped"

if ($relevant.Count -gt 0 -and $matchCount -gt 0) {
    Write-Host "`nVERDICT: Audit trail CAN detect changes in watched folders." -ForegroundColor Green
    Write-Host "These events could replace directory polling for those folders." -ForegroundColor Green
} elseif ($relevant.Count -gt 0 -and $matchCount -eq 0) {
    Write-Host "`nVERDICT: Audit events found but none matched watch folders." -ForegroundColor Yellow
    Write-Host "This may mean no activity in watched folders during the lookback window." -ForegroundColor Yellow
    Write-Host "Try increasing -Hours or triggering a check-in to a watched folder." -ForegroundColor Yellow
} else {
    Write-Host "`nVERDICT: No QC-relevant audit events in the lookback window." -ForegroundColor Yellow
    Write-Host "Try increasing -Hours or check if audit trail is enabled for documents." -ForegroundColor Yellow
}
