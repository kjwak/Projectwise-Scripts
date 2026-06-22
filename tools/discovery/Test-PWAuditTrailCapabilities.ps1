<#
.SYNOPSIS
Read-only discovery of ProjectWise audit trail capabilities for the QC pipeline.

.DESCRIPTION
Connects to ProjectWise using appsettings.json credentials, then probes:
  1. Select-PWSQL availability
  2. Audit trail settings (retention, secondary table)
  3. Recent audit records (sample)
  4. Action type distribution (last 24h)
  5. DOCUMENT_MODIFY events (description changes)
  6. DOCUMENT_ATTR events (attribute changes)
  7. DOCUMENT_CIN events (check-ins)
  8. DOCUMENT_STATE events (workflow state changes)
  9. Audit trail cmdlet availability and parameter shapes

Does not modify any data.

.PARAMETER AppSettingsPath
Path to appsettings.json. Defaults to repo root.

.PARAMETER Hours
Lookback window for distribution queries. Default 24.

.PARAMETER Pretty
Emit formatted output instead of compressed JSON.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [int]$Hours = 24,
    [switch]$Pretty
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Connection.psm1') -Force

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

function _RunSQL {
    param([string]$Sql, [string]$Label)
    try {
        $r = Select-PWSQL -SQLSelectStatement $Sql -ErrorAction Stop
        return @{ success = $true; label = $Label; rowCount = $r.Rows.Count; data = $r }
    } catch {
        return @{ success = $false; label = $Label; error = $_.Exception.Message; data = $null }
    }
}

function _RowsToList {
    param([object]$DataTable)
    if (-not $DataTable -or -not $DataTable.Rows -or $DataTable.Rows.Count -eq 0) { return @() }
    $list = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $val = $row[$col.ColumnName]
            if ($val -is [DBNull]) { $val = $null }
            elseif ($val -is [DateTime]) { $val = Format-QCTimestamp (ConvertTo-QCTimestamp $val) }
            $obj[$col.ColumnName] = $val
        }
        $list.Add([pscustomobject]$obj)
    }
    return @($list)
}

# -- Connect ---------------------------------------------------------------

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw "Cannot load appsettings: $($cfgRes.Message)" }
$config = [hashtable]$cfgRes.Data.config
$pwCfg = ConvertTo-HashtableDeep -Value $config.projectWise

$ds = [string]$pwCfg.datasourceName
$credPath = if ($pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }

Write-Host "Connecting to ProjectWise: $ds ..." -ForegroundColor Cyan
$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw "Credential error: $($credRes.Message)" }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw "Connection error: $($connRes.Message)" }
Write-Host "Connected as $($credRes.Data.userName)." -ForegroundColor Green

$warnings = [System.Collections.Generic.List[string]]::new()
$since = (Get-Date).AddHours(-$Hours).ToString('yyyy-MM-dd HH:mm:ss')

# -- 1. Select-PWSQL availability ------------------------------------------

$sqlCmd = Get-Command -Name Select-PWSQL -ErrorAction SilentlyContinue
$selectPwSqlAvailable = [bool]$sqlCmd
Write-Host "`n[1] Select-PWSQL available: $selectPwSqlAvailable" -ForegroundColor Yellow

if (-not $selectPwSqlAvailable) {
    $warnings.Add('Select-PWSQL cmdlet is not available. Audit trail polling via direct SQL will not work.')
    Write-Host "    BLOCKED: cannot proceed with SQL-based tests." -ForegroundColor Red
}

# -- 2. Audit trail settings -----------------------------------------------

Write-Host "`n[2] Audit trail settings" -ForegroundColor Yellow
$auditSettings = $null
$auditSettingsRaw = $null
try {
    $auditSettingsRaw = Get-PWAuditTrailSettings -ErrorAction Stop
    $auditSettings = $auditSettingsRaw
    Write-Host "    Settings retrieved successfully." -ForegroundColor Green
} catch {
    $warnings.Add("Get-PWAuditTrailSettings failed: $($_.Exception.Message)")
    Write-Host "    Failed: $($_.Exception.Message)" -ForegroundColor Red
}

# -- 3. Sample recent audit records ----------------------------------------

Write-Host "`n[3] Sample recent audit records (TOP 20)" -ForegroundColor Yellow
$sampleResult = $null
$sampleRows = @()
if ($selectPwSqlAvailable) {
    $sampleResult = _RunSQL -Sql "SELECT TOP 20 o_acttime, o_action, o_objtype, o_objguid, o_parentguid, o_itemname, o_itemdesc FROM dms_audt ORDER BY o_acttime DESC" -Label 'recent_sample'
    if ($sampleResult.success) {
        $sampleRows = _RowsToList -DataTable $sampleResult.data
        foreach ($row in $sampleRows) {
            $actionName = $AuditActionMap[[int]$row.o_action]
            if (-not $actionName) { $actionName = "UNKNOWN_$($row.o_action)" }
            $row | Add-Member -NotePropertyName 'actionName' -NotePropertyValue $actionName -Force
        }
        Write-Host "    $($sampleRows.Count) records returned." -ForegroundColor Green
        $sampleRows | Format-Table -Property o_acttime, actionName, o_objtype, o_itemname -AutoSize
    } else {
        $warnings.Add("Recent sample query failed: $($sampleResult.error)")
        Write-Host "    Failed: $($sampleResult.error)" -ForegroundColor Red
    }
}

# -- 4. Action distribution (last N hours) ---------------------------------

Write-Host "[4] Action distribution (last $Hours hours)" -ForegroundColor Yellow
$distRows = @()
if ($selectPwSqlAvailable) {
    $distResult = _RunSQL -Sql "SELECT o_action, COUNT(*) AS cnt FROM dms_audt WHERE o_acttime > '$since' GROUP BY o_action ORDER BY cnt DESC" -Label 'action_distribution'
    if ($distResult.success) {
        $distRows = _RowsToList -DataTable $distResult.data
        foreach ($row in $distRows) {
            $actionName = $AuditActionMap[[int]$row.o_action]
            if (-not $actionName) { $actionName = "UNKNOWN_$($row.o_action)" }
            $row | Add-Member -NotePropertyName 'actionName' -NotePropertyValue $actionName -Force
        }
        Write-Host "    $($distRows.Count) distinct action types." -ForegroundColor Green
        $distRows | Format-Table -Property o_action, actionName, cnt -AutoSize
    } else {
        $warnings.Add("Distribution query failed: $($distResult.error)")
        Write-Host "    Failed: $($distResult.error)" -ForegroundColor Red
    }
}

# -- 5. DOCUMENT_MODIFY events (1002) --------------------------------------

Write-Host "[5] DOCUMENT_MODIFY events (1002) - property/description changes" -ForegroundColor Yellow
$modRows = @()
if ($selectPwSqlAvailable) {
    $modResult = _RunSQL -Sql "SELECT TOP 10 o_acttime, o_objguid, o_itemname, o_itemdesc, o_textparam, o_parentguid FROM dms_audt WHERE o_action = 1002 AND o_acttime > '$since' ORDER BY o_acttime DESC" -Label 'document_modify'
    if ($modResult.success) {
        $modRows = _RowsToList -DataTable $modResult.data
        Write-Host "    $($modRows.Count) DOCUMENT_MODIFY records." -ForegroundColor Green
        $modRows | Format-Table -Property o_acttime, o_itemname, o_itemdesc -AutoSize
    } else {
        $warnings.Add("DOCUMENT_MODIFY query failed: $($modResult.error)")
        Write-Host "    Failed: $($modResult.error)" -ForegroundColor Red
    }
}

# -- 6. DOCUMENT_ATTR events (1003) ----------------------------------------

Write-Host "[6] DOCUMENT_ATTR events (1003) - attribute changes" -ForegroundColor Yellow
$attrRows = @()
if ($selectPwSqlAvailable) {
    $attrResult = _RunSQL -Sql "SELECT TOP 10 o_acttime, o_objguid, o_itemname, o_textparam, o_parentguid FROM dms_audt WHERE o_action = 1003 AND o_acttime > '$since' ORDER BY o_acttime DESC" -Label 'document_attr'
    if ($attrResult.success) {
        $attrRows = _RowsToList -DataTable $attrResult.data
        Write-Host "    $($attrRows.Count) DOCUMENT_ATTR records." -ForegroundColor Green
        if ($attrRows.Count -gt 0) {
            $attrRows | Format-Table -Property o_acttime, o_itemname, o_textparam -AutoSize
        }
    } else {
        $warnings.Add("DOCUMENT_ATTR query failed: $($attrResult.error)")
        Write-Host "    Failed: $($attrResult.error)" -ForegroundColor Red
    }
}

# -- 7. DOCUMENT_CIN events (1007) -----------------------------------------

Write-Host "[7] DOCUMENT_CIN events (1007) - check-ins" -ForegroundColor Yellow
$cinRows = @()
if ($selectPwSqlAvailable) {
    $cinResult = _RunSQL -Sql "SELECT TOP 10 o_acttime, o_objguid, o_itemname, o_parentguid, o_userno, o_comments FROM dms_audt WHERE o_action = 1007 AND o_acttime > '$since' ORDER BY o_acttime DESC" -Label 'document_cin'
    if ($cinResult.success) {
        $cinRows = _RowsToList -DataTable $cinResult.data
        Write-Host "    $($cinRows.Count) DOCUMENT_CIN records." -ForegroundColor Green
        if ($cinRows.Count -gt 0) {
            $cinRows | Format-Table -Property o_acttime, o_itemname, o_parentguid -AutoSize
        }
    } else {
        $warnings.Add("DOCUMENT_CIN query failed: $($cinResult.error)")
        Write-Host "    Failed: $($cinResult.error)" -ForegroundColor Red
    }
}

# -- 8. DOCUMENT_STATE events (1012) ---------------------------------------

Write-Host "[8] DOCUMENT_STATE events (1012) - workflow state changes" -ForegroundColor Yellow
$stateRows = @()
if ($selectPwSqlAvailable) {
    $stateResult = _RunSQL -Sql "SELECT TOP 10 o_acttime, o_objguid, o_itemname, o_textparam, o_numparam1, o_numparam2, o_parentguid FROM dms_audt WHERE o_action = 1012 AND o_acttime > '$since' ORDER BY o_acttime DESC" -Label 'document_state'
    if ($stateResult.success) {
        $stateRows = _RowsToList -DataTable $stateResult.data
        Write-Host "    $($stateRows.Count) DOCUMENT_STATE records." -ForegroundColor Green
        if ($stateRows.Count -gt 0) {
            $stateRows | Format-Table -Property o_acttime, o_itemname, o_textparam, o_numparam1 -AutoSize
        }
    } else {
        $warnings.Add("DOCUMENT_STATE query failed: $($stateResult.error)")
        Write-Host "    Failed: $($stateResult.error)" -ForegroundColor Red
    }
}

# -- 9. Audit trail cmdlet shapes ------------------------------------------

Write-Host "`n[9] Audit trail cmdlet availability" -ForegroundColor Yellow
$auditCmdletNames = @(
    'Get-PWDocumentAuditTrailRecords',
    'Get-PWFolderAuditTrailRecords',
    'Get-PWUserAuditTrailRecords',
    'Get-PWAuditTrailRecordsFromPreviousDays',
    'Export-PWAuditTrailToSQLite',
    'Export-PWAuditTrailNotificationReport',
    'Get-PWAuditTrailSettings',
    'Set-PWAuditTrailSetting',
    'Set-PWAuditTrailSettingsAll',
    'Clear-PWAuditTrailSetting',
    'Clear-PWAuditTrailSettingsAll',
    'Move-PWLogInOutAuditTrailRecords',
    'Select-PWSQL'
)

$auditCmdlets = [System.Collections.Generic.List[pscustomobject]]::new()
foreach ($name in $auditCmdletNames) {
    $cmd = Get-Command -Name $name -ErrorAction SilentlyContinue
    $entry = [ordered]@{
        name      = $name
        available = [bool]$cmd
        type      = if ($cmd) { [string]$cmd.CommandType } else { $null }
        module    = if ($cmd) { [string]$cmd.ModuleName } else { $null }
    }
    if ($cmd -and $cmd.Parameters) {
        $paramNames = @($cmd.Parameters.Keys | Where-Object { $_ -notmatch '^(Verbose|Debug|ErrorAction|WarningAction|InformationAction|ErrorVariable|WarningVariable|InformationVariable|OutVariable|OutBuffer|PipelineVariable)$' } | Sort-Object)
        $entry['parameters'] = $paramNames
    }
    $auditCmdlets.Add([pscustomobject]$entry)
    $status = if ($entry.available) { 'YES' } else { 'NO ' }
    Write-Host "    [$status] $name" -ForegroundColor $(if ($entry.available) { 'Green' } else { 'DarkGray' })
}

# -- 10. Watermark feasibility ---------------------------------------------

Write-Host "`n[10] Watermark feasibility - timestamp ordering check" -ForegroundColor Yellow
$watermarkCheck = $null
if ($selectPwSqlAvailable) {
    $wmResult = _RunSQL -Sql "SELECT TOP 100 o_acttime FROM dms_audt ORDER BY o_acttime DESC" -Label 'watermark_check'
    if ($wmResult.success -and $wmResult.data.Rows.Count -gt 1) {
        $times = @($wmResult.data.Rows | ForEach-Object { Format-QCTimestamp (ConvertTo-QCTimestamp $_.o_acttime) })
        $duplicateTimestamps = ($times | Group-Object | Where-Object { $_.Count -gt 1 }).Count
        $oldestInBatch = $times[-1]
        $newestInBatch = $times[0]
        $watermarkCheck = [ordered]@{
            rowsChecked          = $times.Count
            newestTimestamp       = $newestInBatch
            oldestTimestamp       = $oldestInBatch
            duplicateTimestamps  = $duplicateTimestamps
            duplicateRatio       = [math]::Round($duplicateTimestamps / [math]::Max(1, $times.Count), 3)
        }
        Write-Host "    Checked $($times.Count) records. Duplicate timestamps: $duplicateTimestamps ($([math]::Round($duplicateTimestamps / [math]::Max(1, $times.Count) * 100, 1))%)" -ForegroundColor Green
        Write-Host "    Range: $oldestInBatch -> $newestInBatch" -ForegroundColor Green
        if ($duplicateTimestamps -gt 0) {
            Write-Host "    NOTE: Duplicate timestamps exist. Watermark polling should use >= with deduplication or overlap by a few seconds." -ForegroundColor DarkYellow
        }
    } elseif ($wmResult.success) {
        Write-Host "    Not enough records to evaluate." -ForegroundColor DarkYellow
    } else {
        $warnings.Add("Watermark check failed: $($wmResult.error)")
        Write-Host "    Failed: $($wmResult.error)" -ForegroundColor Red
    }
}

# -- 11. Table row count estimate ------------------------------------------

Write-Host "`n[11] dms_audt table size" -ForegroundColor Yellow
$tableSize = $null
if ($selectPwSqlAvailable) {
    $countResult = _RunSQL -Sql "SELECT COUNT(*) AS total_rows FROM dms_audt" -Label 'table_count'
    if ($countResult.success) {
        $totalRows = $countResult.data.Rows[0].total_rows
        $tableSize = [ordered]@{ totalRows = $totalRows }
        Write-Host "    Total rows in dms_audt: $totalRows" -ForegroundColor Green
    } else {
        $warnings.Add("Table count query failed: $($countResult.error)")
        Write-Host "    Failed: $($countResult.error)" -ForegroundColor Red
    }
}

# -- Disconnect ------------------------------------------------------------

Disconnect-PW | Out-Null

# -- Build result ----------------------------------------------------------

$_recentSample   = if ($sampleRows -and $sampleRows.Count -gt 0) { ,@($sampleRows | Select-Object -First 5) } else { ,@() }
$_actionDist     = if ($distRows -and $distRows.Count -gt 0) { ,@($distRows) } else { ,@() }
$_modSample      = if ($modRows -and $modRows.Count -gt 0) { ,@($modRows) } else { ,@() }
$_attrSample     = if ($attrRows -and $attrRows.Count -gt 0) { ,@($attrRows) } else { ,@() }
$_cinSample      = if ($cinRows -and $cinRows.Count -gt 0) { ,@($cinRows) } else { ,@() }
$_stateSample    = if ($stateRows -and $stateRows.Count -gt 0) { ,@($stateRows) } else { ,@() }

$result = [ordered]@{
    timestamp               = Get-QCTimestamp
    timezone                = 'MST (UTC-7)'
    datasource              = $ds
    selectPwSqlAvailable    = $selectPwSqlAvailable
    auditSettings           = $auditSettingsRaw
    tableSize               = $tableSize
    watermarkFeasibility    = $watermarkCheck
    recentSample            = $_recentSample
    actionDistribution      = $_actionDist
    documentModifySample    = $_modSample
    documentAttrSample      = $_attrSample
    documentCheckinSample   = $_cinSample
    documentStateSample     = $_stateSample
    auditCmdlets            = @($auditCmdlets.ToArray())
    warnings                = @($warnings)
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Select-PWSQL:        $selectPwSqlAvailable"
Write-Host "Audit settings:      $(if ($auditSettingsRaw) { 'Retrieved' } else { 'FAILED' })"
Write-Host "Table rows:          $(if ($tableSize) { $tableSize.totalRows } else { 'N/A' })"
Write-Host "Action types (${Hours}h): $($distRows.Count)"
Write-Host "Warnings:            $($warnings.Count)"

if ($warnings.Count -gt 0) {
    Write-Host "`nWarnings:" -ForegroundColor DarkYellow
    foreach ($w in $warnings) { Write-Host "  - $w" -ForegroundColor DarkYellow }
}

$depth = 30
if ($Pretty) { $result | ConvertTo-Json -Depth $depth }
else { $result | ConvertTo-Json -Depth $depth -Compress }
