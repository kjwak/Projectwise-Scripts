<#
.SYNOPSIS
Show parameters and output fields exposed by ProjectWise audit events (dms_audt + cmdlets).

.DESCRIPTION
Read-only discovery script for audit-trail filtering design:
  1. dms_audt column schema (all o_* fields)
  2. Population stats for o_parentguid / o_guidparam by action and object type
  3. Sample rows with every column for QC-relevant actions
  4. Cmdlet output property names (Get-PWDocumentAuditTrailRecords, Get-PWFolderAuditTrailRecords)
  5. Watch-folder GUID cache vs parent GUID overlap (filter feasibility)

Does not modify ProjectWise or the QC database.

.PARAMETER AppSettingsPath
Path to appsettings.json. Defaults to repo root.

.PARAMETER Hours
Lookback window for stats and samples. Default 24.

.PARAMETER SamplePerAction
Max sample rows per QC-relevant action code. Default 2.

.PARAMETER Pretty
Emit formatted JSON at the end.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [int]$Hours = 24,
    [int]$SamplePerAction = 2,
    [switch]$Pretty
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

foreach ($mod in @('Core.Results.psm1', 'Core.Runtime.psm1', 'PW.Connection.psm1', 'PW.AuditPoller.psm1')) {
    Import-Module (Join-Path $repoRoot "modules\$mod") -Force -WarningAction SilentlyContinue
}

$AuditActionMap = @{
    1    = 'FOLDER_CREATE'
    2    = 'FOLDER_MODIFY'
    3    = 'FOLDER_WFLOW'
    4    = 'FOLDER_DELETE'
    5    = 'FOLDER_STATE'
    1001 = 'DOCUMENT_CREATE'
    1002 = 'DOCUMENT_MODIFY'
    1003 = 'DOCUMENT_ATTR'
    1004 = 'DOCUMENT_FILE_ADD'
    1005 = 'DOCUMENT_FILE_REM'
    1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN'
    1008 = 'DOCUMENT_VIEW'
    1009 = 'DOCUMENT_CHOUT'
    1010 = 'DOCUMENT_CPOUT'
    1011 = 'DOCUMENT_GOUT'
    1012 = 'DOCUMENT_STATE'
    1013 = 'DOCUMENT_FINAL_S'
    1014 = 'DOCUMENT_FINAL_R'
    1015 = 'DOCUMENT_VERSION'
    1016 = 'DOCUMENT_MOVE'
    1020 = 'DOCUMENT_DELETE'
    1022 = 'DOCUMENT_FREE'
    1027 = 'DOCUMENT_IMPORT'
    3001 = 'USER_LOGIN'
    3002 = 'USER_LOGOUT'
}

$QCRelevantActions = @(1001, 1002, 1003, 1006, 1007, 1012, 1015, 1020)

$DmsAudtColumnDocs = [ordered]@{
    o_objtype   = 'Object type (1=Folder, 2=Document, 3=Set, etc.)'
    o_objguid   = 'GUID of the affected object (document or folder)'
    o_objno     = 'Numeric ID of the affected object'
    o_action    = 'Action type numeric code (see action map)'
    o_acttime   = 'Timestamp of the action (UTC wall clock in PW)'
    o_userno    = 'User number who performed the action'
    o_comments  = 'Optional comment text'
    o_numparam1 = 'Numeric parameter 1 (action-specific, e.g. state IDs for DOCUMENT_STATE)'
    o_numparam2 = 'Numeric parameter 2 (action-specific)'
    o_textparam = 'Text parameter (attribute name, state name, etc.)'
    o_guidparam = 'GUID parameter (secondary object GUID when applicable)'
    o_itemname  = 'Name of the affected item'
    o_itemdesc  = 'Description of the affected item'
    o_parentguid = 'Parent GUID - for document events this is the containing folder GUID (use for folder filtering)'
}

function _RunSQL {
    param([string]$Sql, [string]$Label)
    try {
        $r = Select-PWSQL -SQLSelectStatement $Sql -ErrorAction Stop
        return @{ success = $true; label = $Label; data = $r }
    } catch {
        return @{ success = $false; label = $Label; error = $_.Exception.Message; data = $null }
    }
}

function _SqlRows {
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

function _RowToOrderedHash {
    param($Row)
    if ($null -eq $Row) { return $null }
    $obj = [ordered]@{}
    if ($Row -is [System.Data.DataRow]) {
        foreach ($col in $Row.Table.Columns) {
            $val = $Row[$col.ColumnName]
            if ($val -is [DBNull]) { $val = $null }
            elseif ($val -is [DateTime]) { $val = $val.ToString('yyyy-MM-dd HH:mm:ss') }
            $obj[$col.ColumnName] = $val
        }
        return $obj
    }
    foreach ($prop in $Row.PSObject.Properties) {
        $val = $prop.Value
        if ($val -is [DBNull]) { $val = $null }
        elseif ($val -is [DateTime]) { $val = $val.ToString('yyyy-MM-dd HH:mm:ss') }
        $obj[$prop.Name] = $val
    }
    return $obj
}

function _NormalizeGuidKey {
    param([string]$Guid)
    if ([string]::IsNullOrWhiteSpace($Guid)) { return $null }
    return ($Guid.Trim().Trim('{}').ToLowerInvariant())
}

function _DescribeObjectProperties {
    param([object]$Obj)
    if ($null -eq $Obj) { return @() }
    $names = @()
    try { $names = @($Obj.PSObject.Properties.Name | Sort-Object) } catch { }
    return $names
}

function _ActionName([int]$Code) {
    if ($AuditActionMap.ContainsKey($Code)) { return $AuditActionMap[$Code] }
    return "UNKNOWN_$Code"
}

function _ObjTypeLabel([int]$Type) {
    switch ($Type) {
        1 { return 'Folder' }
        2 { return 'Document' }
        3 { return 'Set' }
        default { return "Type_$Type" }
    }
}

function _IsPopulatedGuid([object]$Value) {
    if ($null -eq $Value -or $Value -is [DBNull]) { return $false }
    $s = ([string]$Value).Trim().Trim('{}')
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }
    if ($s -eq '00000000-0000-0000-0000-000000000000') { return $false }
    return $true
}

# -- Load config & connect ---------------------------------------------------

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw "Cannot load appsettings: $($cfgRes.Message)" }
$config = [hashtable]$cfgRes.Data.config
$pwCfg = ConvertTo-HashtableDeep -Value $config.projectWise

$ds = [string]$pwCfg.datasourceName
$credPath = if ($pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }

Write-Host "=== ProjectWise audit event schema ===" -ForegroundColor Cyan
Write-Host "Datasource: $ds"
Write-Host "Lookback:   $Hours hours`n"

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw "Credential error: $($credRes.Message)" }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw "Connection error: $($connRes.Message)" }

$pingResult = _RunSQL -Sql 'SELECT TOP 1 o_acttime FROM dms_audt ORDER BY o_acttime DESC' -Label 'connection_ping'
if (-not $pingResult.success) {
    throw "Connected but Select-PWSQL failed: $($pingResult.error). Run from ProjectWise PowerShell (MTA) with valid credentials."
}
Write-Host "Connected as $($credRes.Data.userName). dms_audt reachable.`n" -ForegroundColor Green

$since = (Get-Date).ToUniversalTime().AddHours(-$Hours).ToString('yyyy-MM-dd HH:mm:ss')
$warnings = [System.Collections.Generic.List[string]]::new()

# -- 1. Documented dms_audt columns ------------------------------------------

Write-Host "[1] dms_audt columns (documented + observed)" -ForegroundColor Yellow
Write-Host "    NOTE: There is no separate 'folder-guid' column. Use o_parentguid for the containing folder on document events.`n" -ForegroundColor DarkGray

$schemaRows = @()
$schemaResult = _RunSQL -Sql "SELECT TOP 1 * FROM dms_audt ORDER BY o_acttime DESC" -Label 'schema_probe'
if ($schemaResult.success) {
    $probeRows = _SqlRows $schemaResult.data
    if ($probeRows.Count -gt 0 -and $probeRows[0] -is [System.Data.DataRow]) {
        foreach ($col in $probeRows[0].Table.Columns) {
            $name = [string]$col.ColumnName
            $desc = if ($DmsAudtColumnDocs.Contains($name)) { $DmsAudtColumnDocs[$name] } else { '(observed; not in QC docs)' }
            $schemaRows += [pscustomobject]@{ column = $name; dataType = [string]$col.DataType.Name; description = $desc }
        }
    }
}
if ($schemaRows.Count -eq 0) {
    foreach ($name in $DmsAudtColumnDocs.Keys) {
        $schemaRows += [pscustomobject]@{ column = $name; dataType = '(unknown)'; description = $DmsAudtColumnDocs[$name] }
    }
    $warnings.Add('Could not introspect dms_audt columns from live data; showing documented schema only.')
}

$schemaRows | Format-Table -Property column, dataType, description -Wrap -AutoSize

$parentGuidDoc = $schemaRows | Where-Object { $_.column -eq 'o_parentguid' }
if ($parentGuidDoc) {
    Write-Host "    o_parentguid present in schema: YES" -ForegroundColor Green
} else {
    Write-Host "    o_parentguid present in schema: NO (unexpected)" -ForegroundColor Red
    $warnings.Add('o_parentguid column not found in dms_audt introspection.')
}

# -- 2. Population stats for GUID fields ---------------------------------------

Write-Host "[2] GUID field population (last $Hours h)" -ForegroundColor Yellow

$guidStats = @()
$statsSql = @"
SELECT
    o_action,
    o_objtype,
    COUNT(*) AS total,
    SUM(CASE WHEN o_parentguid IS NOT NULL AND CAST(o_parentguid AS varchar(50)) NOT IN ('', '00000000-0000-0000-0000-000000000000') THEN 1 ELSE 0 END) AS with_parentguid,
    SUM(CASE WHEN o_guidparam IS NOT NULL AND CAST(o_guidparam AS varchar(50)) NOT IN ('', '00000000-0000-0000-0000-000000000000') THEN 1 ELSE 0 END) AS with_guidparam
FROM dms_audt
WHERE o_acttime >= '$since'
GROUP BY o_action, o_objtype
ORDER BY total DESC
"@

$statsResult = _RunSQL -Sql $statsSql -Label 'guid_population'
if ($statsResult.success) {
    foreach ($row in (_SqlRows $statsResult.data)) {
        $total = [int]$row.total
        $withParent = [int]$row.with_parentguid
        $withGuidParam = [int]$row.with_guidparam
        $action = [int]$row.o_action
        $objtype = [int]$row.o_objtype
        $guidStats += [pscustomobject]@{
            action           = $action
            actionName       = _ActionName $action
            objType          = $objtype
            objTypeLabel     = _ObjTypeLabel $objtype
            total            = $total
            withParentGuid   = $withParent
            parentGuidPct    = if ($total -gt 0) { [math]::Round(100.0 * $withParent / $total, 1) } else { 0 }
            withGuidParam    = $withGuidParam
            guidParamPct     = if ($total -gt 0) { [math]::Round(100.0 * $withGuidParam / $total, 1) } else { 0 }
        }
    }
    $guidStats | Format-Table -Property actionName, objTypeLabel, total, withParentGuid, parentGuidPct, withGuidParam, guidParamPct -AutoSize

    $docQc = @($guidStats | Where-Object { $_.objType -eq 2 -and ($QCRelevantActions -contains $_.action) })
    if ($docQc.Count -gt 0) {
        $docTotal = ($docQc | Measure-Object -Property total -Sum).Sum
        $docWithParent = ($docQc | Measure-Object -Property withParentGuid -Sum).Sum
        $pct = if ($docTotal -gt 0) { [math]::Round(100.0 * $docWithParent / $docTotal, 1) } else { 0 }
        Write-Host "    QC-relevant document events: $docWithParent / $docTotal have o_parentguid (${pct} pct)" -ForegroundColor $(if ($pct -ge 95) { 'Green' } elseif ($pct -ge 50) { 'Yellow' } else { 'Red' })
    }
} else {
    $warnings.Add("GUID population query failed: $($statsResult.error)")
    Write-Host "    Failed: $($statsResult.error)" -ForegroundColor Red
}

# -- 3. Sample rows per QC-relevant action ------------------------------------

Write-Host "`n[3] Sample audit rows (all columns) for QC-relevant actions" -ForegroundColor Yellow

$actionSamples = [ordered]@{}
foreach ($actionCode in $QCRelevantActions) {
    $actionName = _ActionName $actionCode
    $sampleSql = @"
SELECT TOP $SamplePerAction *
FROM dms_audt
WHERE o_action = $actionCode AND o_acttime >= '$since'
ORDER BY o_acttime DESC
"@
    $sampleResult = _RunSQL -Sql $sampleSql -Label "sample_$actionName"
    $rows = @()
    if ($sampleResult.success) {
        foreach ($r in (_SqlRows $sampleResult.data)) {
            $h = _RowToOrderedHash $r
            if ($h) {
                $h['actionName'] = $actionName
                $h['objTypeLabel'] = _ObjTypeLabel ([int]$h['o_objtype'])
                $rows += $h
            }
        }
    } else {
        $warnings.Add("Sample query for $actionName failed: $($sampleResult.error)")
    }
    $actionSamples[$actionName] = $rows

    Write-Host "`n  --- $actionName ($actionCode) - $($rows.Count) samples ---" -ForegroundColor Cyan
    if ($rows.Count -eq 0) {
        Write-Host "    (no events in lookback window)" -ForegroundColor DarkGray
        continue
    }
    foreach ($sample in $rows) {
        $pg = $sample['o_parentguid']
        $og = $sample['o_objguid']
        $parentStatus = if (_IsPopulatedGuid $pg) { 'POPULATED' } else { 'EMPTY' }
        Write-Host ("    {0}  doc={1}  parent={2}  [{3}]" -f $sample['o_acttime'], $sample['o_itemname'], $pg, $parentStatus)
        Write-Host ("      o_textparam={0}  o_numparam1={1}  o_numparam2={2}" -f $sample['o_textparam'], $sample['o_numparam1'], $sample['o_numparam2'])
    }
}

# -- 4. Cmdlet output shapes ---------------------------------------------------

Write-Host "`n[4] Audit trail cmdlet parameters and output properties" -ForegroundColor Yellow

$auditCmdlets = @(
    'Get-PWDocumentAuditTrailRecords',
    'Get-PWFolderAuditTrailRecords',
    'Get-PWUserAuditTrailRecords',
    'Get-PWAuditTrailRecordsFromPreviousDays'
)

$cmdletInfo = [System.Collections.Generic.List[pscustomobject]]::new()
$cmdletSampleProps = [ordered]@{}

foreach ($name in $auditCmdlets) {
    $cmd = Get-Command -Name $name -ErrorAction SilentlyContinue
    $entry = [ordered]@{
        name      = $name
        available = [bool]$cmd
        parameters = @()
        outputProperties = @()
        sampleNote = $null
    }
    if ($cmd) {
        $entry.parameters = @($cmd.Parameters.Keys |
            Where-Object { $_ -notmatch '^(Verbose|Debug|ErrorAction|WarningAction|InformationAction|ErrorVariable|WarningVariable|InformationVariable|OutVariable|OutBuffer|PipelineVariable)$' } |
            Sort-Object)
        Write-Host "    [$name]" -ForegroundColor Green
        Write-Host ("      Parameters: {0}" -f ($entry.parameters -join ', '))
    } else {
        Write-Host "    [$name] not available" -ForegroundColor DarkGray
    }
    $cmdletInfo.Add([pscustomobject]$entry)
}

# Try to fetch one cmdlet sample for property introspection
$docSampleProps = @()
$folderSampleProps = @()
try {
    $recentDoc = _RunSQL -Sql "SELECT TOP 1 o_objguid FROM dms_audt WHERE o_objtype = 2 AND o_acttime >= '$since' ORDER BY o_acttime DESC" -Label 'doc_guid_for_cmdlet'
    if ($recentDoc.success) {
        $docRows = _SqlRows $recentDoc.data
        if ($docRows.Count -gt 0) {
            $dg = [string]$docRows[0].o_objguid
            if ((_IsPopulatedGuid $dg) -and (Get-Command -Name 'Get-PWDocumentAuditTrailRecords' -ErrorAction SilentlyContinue) -and (Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue)) {
                $doc = Get-PWDocumentsByGUIDs -DocumentGUIDs @($dg) -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($doc) {
                    $recs = @(Get-PWDocumentAuditTrailRecords -InputDocuments @($doc) -ErrorAction SilentlyContinue)
                    if ($recs.Count -gt 0) {
                        $docSampleProps = _DescribeObjectProperties $recs[0]
                        Write-Host "`n    Get-PWDocumentAuditTrailRecords output properties (first record):" -ForegroundColor Green
                        Write-Host ("      {0}" -f ($docSampleProps -join ', '))
                        $hasParent = @($docSampleProps | Where-Object { $_ -match 'parent|folder|guid' -and $_ -notmatch 'objguid' })
                        if ($hasParent.Count -gt 0) {
                            Write-Host ("      GUID/parent-related: {0}" -f ($hasParent -join ', ')) -ForegroundColor Cyan
                        }
                    }
                }
            }
        }
    }
} catch {
    $warnings.Add("Get-PWDocumentAuditTrailRecords sample failed: $($_.Exception.Message)")
}

try {
    $recentFolder = _RunSQL -Sql "SELECT TOP 1 o_objguid FROM dms_audt WHERE o_objtype = 1 AND o_acttime >= '$since' ORDER BY o_acttime DESC" -Label 'folder_guid_for_cmdlet'
    if ($recentFolder.success) {
        $folderRows = _SqlRows $recentFolder.data
        if ($folderRows.Count -gt 0) {
            $fg = [string]$folderRows[0].o_objguid
            if (($fg) -and (Get-Command -Name 'Get-PWFolderAuditTrailRecords' -ErrorAction SilentlyContinue) -and (Get-Command -Name 'Get-PWFoldersByGUIDs' -ErrorAction SilentlyContinue)) {
                $folder = Get-PWFoldersByGUIDs -FolderGUIDs @($fg) -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($folder) {
                    $recs = @(Get-PWFolderAuditTrailRecords -InputFolder $folder -ErrorAction SilentlyContinue)
                    if ($recs.Count -gt 0) {
                        $folderSampleProps = _DescribeObjectProperties $recs[0]
                        Write-Host "`n    Get-PWFolderAuditTrailRecords output properties (first record):" -ForegroundColor Green
                        Write-Host ("      {0}" -f ($folderSampleProps -join ', '))
                        $hasParent = @($folderSampleProps | Where-Object { $_ -match 'parent|folder|guid' })
                        if ($hasParent.Count -gt 0) {
                            Write-Host ("      GUID/parent-related: {0}" -f ($hasParent -join ', ')) -ForegroundColor Cyan
                        }
                    }
                }
            }
        }
    }
} catch {
    $warnings.Add("Get-PWFolderAuditTrailRecords sample failed: $($_.Exception.Message)")
}

$cmdletSampleProps['Get-PWDocumentAuditTrailRecords'] = $docSampleProps
$cmdletSampleProps['Get-PWFolderAuditTrailRecords'] = $folderSampleProps

# -- 5. Watch-folder GUID filter feasibility -----------------------------------

Write-Host "`n[5] Watch-folder GUID filter feasibility" -ForegroundColor Yellow

$watchGuidSet = @{}
$watchFolders = @()
try {
    $warmRes = Sync-AuditPollerWatchFolderGuidCache -Config $config
    if ($warmRes.IsSuccess -and $warmRes.Data) {
        Write-Host ("    Sync-AuditPollerWatchFolderGuidCache: warmed {0} folder(s)" -f $warmRes.Data.warmed) -ForegroundColor Green
    }
} catch {
    $warnings.Add("Watch folder GUID cache warm failed: $($_.Exception.Message)")
}

if (Get-Command -Name 'Get-QCPwFolderGuidCache' -ErrorAction SilentlyContinue) {
    try {
        $cacheRes = Get-QCPwFolderGuidCache -Config $config
        if ($cacheRes.IsSuccess -and $cacheRes.Data.cache) {
            foreach ($k in $cacheRes.Data.cache.Keys) {
                $watchGuidSet[(_NormalizeGuidKey $k)] = $cacheRes.Data.cache[$k]
            }
        }
    } catch { }
}

foreach ($k in @($watchGuidSet.Keys | Sort-Object)) {
    $watchFolders += [pscustomobject]@{ folderGuid = $k; folderPath = $watchGuidSet[$k] }
}

Write-Host ("    Watch folder GUIDs cached: {0}" -f $watchFolders.Count)
if ($watchFolders.Count -gt 0) {
    $watchFolders | Select-Object -First 8 | Format-Table -Property folderGuid, folderPath -AutoSize
}

$filterStats = @{
    qcEventsInWindow = 0
    withParentGuid = 0
    matchingWatchFolder = 0
    matchingWatchFolderPct = 0
    uniqueParentGuids = 0
    uniqueMatchingParentGuids = 0
}

$filterSql = @"
SELECT o_action, o_objtype, o_objguid, o_parentguid, o_itemname, o_acttime
FROM dms_audt
WHERE o_acttime >= '$since'
  AND o_objtype = 2
  AND o_action IN ($(($QCRelevantActions -join ',')))
"@

$filterResult = _RunSQL -Sql $filterSql -Label 'watch_filter_probe'
$matchingSamples = [System.Collections.Generic.List[pscustomobject]]::new()
$parentGuidUniverse = @{}

if ($filterResult.success) {
    foreach ($row in (_SqlRows $filterResult.data)) {
        $filterStats.qcEventsInWindow++
        $pg = _NormalizeGuidKey ([string]$row.o_parentguid)
        if (_IsPopulatedGuid $row.o_parentguid) {
            $filterStats.withParentGuid++
            if ($pg) { $parentGuidUniverse[$pg] = $true }
            if ($watchGuidSet.ContainsKey($pg)) {
                $filterStats.matchingWatchFolder++
                if ($matchingSamples.Count -lt 10) {
                    $matchingSamples.Add([pscustomobject]@{
                        acttime    = [string]$row.o_acttime
                        actionName = _ActionName ([int]$row.o_action)
                        itemname   = [string]$row.o_itemname
                        parentGuid = $pg
                        folderPath = $watchGuidSet[$pg]
                    })
                }
            }
        }
    }
    $filterStats.uniqueParentGuids = $parentGuidUniverse.Count
    $filterStats.uniqueMatchingParentGuids = @($parentGuidUniverse.Keys | Where-Object { $watchGuidSet.ContainsKey($_) }).Count
    if ($filterStats.withParentGuid -gt 0) {
        $filterStats.matchingWatchFolderPct = [math]::Round(100.0 * $filterStats.matchingWatchFolder / $filterStats.withParentGuid, 1)
    }

    Write-Host ("    QC document events in window:     {0}" -f $filterStats.qcEventsInWindow)
    Write-Host ("    With o_parentguid:               {0}" -f $filterStats.withParentGuid)
    $matchPctLabel = "{0} pct of events with parent GUID" -f $filterStats.matchingWatchFolderPct
    Write-Host ("    Matching cached watch folder:    {0} ({1})" -f $filterStats.matchingWatchFolder, $matchPctLabel) -ForegroundColor $(if ($filterStats.matchingWatchFolder -gt 0) { 'Green' } else { 'DarkYellow' })
    Write-Host ("    Unique parent GUIDs:             {0}" -f $filterStats.uniqueParentGuids)
    Write-Host ("    Unique parent GUIDs in watch set: {0}" -f $filterStats.uniqueMatchingParentGuids)

    if ($matchingSamples.Count -gt 0) {
        Write-Host "`n    Sample events that WOULD match o_parentguid filter:" -ForegroundColor Green
        $matchingSamples | Format-Table -Property acttime, actionName, itemname, parentGuid, folderPath -AutoSize
    } elseif ($filterStats.qcEventsInWindow -gt 0 -and $watchFolders.Count -gt 0) {
        Write-Host "    No QC events matched watch folder GUIDs in this window (may be normal if no activity in watched folders)." -ForegroundColor DarkYellow
    }
} else {
    $warnings.Add("Watch filter probe failed: $($filterResult.error)")
    Write-Host "    Failed: $($filterResult.error)" -ForegroundColor Red
}

# -- Summary -------------------------------------------------------------------

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Source of truth for polling:  dms_audt via Select-PWSQL (used by PW.AuditPoller)"
Write-Host "Folder filter field:          o_parentguid (NOT a separate folder-guid column)"
Write-Host "Pipeline mapping:             o_parentguid -> audit_events.pw_parentguid -> resolved_folder"
Write-Host "Recommended filter:           WHERE o_parentguid IN (<watch folder GUIDs>) AND o_action IN (...)"
Write-Host ""
Write-Host "QC-relevant actions polled today:"
foreach ($code in $QCRelevantActions) {
    Write-Host ("  {0,-5} {1}" -f $code, (_ActionName $code))
}

if ($warnings.Count -gt 0) {
    Write-Host "`nWarnings:" -ForegroundColor DarkYellow
    foreach ($w in $warnings) { Write-Host "  - $w" -ForegroundColor DarkYellow }
}

Disconnect-PW | Out-Null

$result = [ordered]@{
    timestamp              = Get-QCTimestamp
    datasource             = $ds
    lookbackHours          = $Hours
    sinceUtc               = $since
    dmsAudtColumns         = @($schemaRows)
    parentGuidField        = 'o_parentguid'
    folderGuidField        = $null
    filterRecommendation   = 'Filter document events by o_parentguid against watch-folder GUID cache (Sync-AuditPollerWatchFolderGuidCache / pw_folder_cache).'
    guidPopulationStats    = @($guidStats)
    actionSamples          = $actionSamples
    auditCmdlets           = @($cmdletInfo.ToArray())
    cmdletOutputProperties = $cmdletSampleProps
    watchFolderGuids       = @($watchFolders)
    watchFilterStats       = $filterStats
    matchingEventSamples   = @($matchingSamples.ToArray())
    warnings               = @($warnings)
}

$depth = 12
if ($Pretty) { $result | ConvertTo-Json -Depth $depth }
else { $result | ConvertTo-Json -Depth $depth -Compress }
