<#
.SYNOPSIS
Direct dms_audt query: show raw o_acttime from ProjectWise vs pipeline clocks.

.PARAMETER Hours
How far back to query (default 48, to catch "yesterday" events appearing today).

.PARAMETER Top
Number of newest rows to print (default 25).

.PARAMETER ItemNameLike
Optional SQL LIKE filter on o_itemname, e.g. '%YourSheet.pdf%'.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [int]$Hours = 48,
    [int]$Top = 25,
    [string]$ItemNameLike = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

foreach ($mod in @('Core.Results.psm1', 'Core.Runtime.psm1', 'PW.Connection.psm1', 'PW.AuditPoller.psm1')) {
    Import-Module (Join-Path $repoRoot "modules\$mod") -Force -WarningAction SilentlyContinue
}

$config = Get-QCAppSettingsConfig -Path $AppSettingsPath
if ($config -isnot [hashtable]) { $config = ConvertTo-HashtableDeep -Value $config }

$displayTz = Get-QCDisplayTimeZone

function _PwActTimeToUtc([DateTime]$dt) {
    if ($dt.Kind -eq [DateTimeKind]::Utc) { return $dt }
    if ($dt.Kind -eq [DateTimeKind]::Local) { return $dt.ToUniversalTime() }
    return [DateTime]::SpecifyKind($dt, [DateTimeKind]::Utc)
}

function _FmtDisplayFromPw([DateTime]$dt) {
    $local = [TimeZoneInfo]::ConvertTimeFromUtc((_PwActTimeToUtc $dt), $displayTz)
    return $local.ToString('yyyy-MM-dd HH:mm:ss') + ' (' + $displayTz.Id + ')'
}

function _FmtPipelinePwActTime([object]$Raw) {
    if ($null -eq $Raw -or $Raw -is [DBNull]) { return $null }
    $tz = $displayTz
    if ($Raw -is [DateTime]) {
        $mt = [TimeZoneInfo]::ConvertTimeFromUtc((_PwActTimeToUtc $Raw), $tz)
        return $mt.ToString('yyyy-MM-dd HH:mm:ss')
    }
    $s = ([string]$Raw).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try {
        $parsed = [DateTime]::Parse($s)
        $mt = [TimeZoneInfo]::ConvertTimeFromUtc((_PwActTimeToUtc $parsed), $tz)
        return $mt.ToString('yyyy-MM-dd HH:mm:ss')
    } catch { return $s }
}

function _DescribeActTime($raw) {
    if ($null -eq $raw -or $raw -is [DBNull]) { return @{ raw = ''; type = 'null' } }
    $typeName = $raw.GetType().FullName
    $out = @{ raw = [string]$raw; type = $typeName }
    if ($raw -is [DateTime]) {
        $out.kind = [string]$raw.Kind
        $utcAssumed = if ($raw.Kind -eq [DateTimeKind]::Unspecified) {
            [DateTime]::SpecifyKind($raw, [DateTimeKind]::Utc)
        } else {
            $raw.ToUniversalTime()
        }
        $out.utcAssumed = $utcAssumed.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC (PW Unspecified treated as UTC)'
        $out.display = _FmtDisplayFromPw $raw
        $out.wrongLocalAsUtc = $raw.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + ' UTC (WRONG: ToUniversalTime on Unspecified uses machine local)'
        $out.toStringLocal = $raw.ToString('yyyy-MM-dd HH:mm:ss') + ' (DateTime.ToString on this machine)'
    }
    return $out
}

# Clocks
$nowLocal = [DateTimeOffset]::Now
$nowUtc = [DateTimeOffset]::UtcNow
$nowDisplay = [TimeZoneInfo]::ConvertTimeFromUtc($nowUtc.UtcDateTime, $displayTz)
Write-Host "=== Clocks (compare to PW o_acttime) ===" -ForegroundColor Cyan
Write-Host ("  Windows local:     {0}" -f $nowLocal.ToString('yyyy-MM-dd HH:mm:ss zzz'))
Write-Host ("  UTC:               {0}" -f $nowUtc.ToString('yyyy-MM-dd HH:mm:ss') + ' Z')
Write-Host ("  Display (config):  {0} ({1})" -f $nowDisplay.ToString('yyyy-MM-dd HH:mm:ss'), $displayTz.Id)
Write-Host ("  Get-QCTimestamp:   {0}" -f (Get-QCTimestamp))

$sinceUtc = $nowUtc.UtcDateTime.AddHours(-$Hours).ToString('yyyy-MM-dd HH:mm:ss')
Write-Host ""
Write-Host "  SQL filter (UTC, matches watcher): o_acttime >= '$sinceUtc'  (${Hours}h lookback)" -ForegroundColor DarkGray

# Poll window (what watcher uses)
$wmPath = Join-Path (Join-Path $config.queue.rootDir '_watcher') 'audit-capture-watermark.txt'
$lookback = 1200
try {
    if ($config.auditPoller.lookbackSeconds) { $lookback = [int]$config.auditPoller.lookbackSeconds }
} catch { }
$poll = Get-AuditTrailPollWindow -Config $config -WatermarkPath $wmPath -LookbackSeconds $lookback
Write-Host "`n=== Watcher poll window (Get-AuditTrailPollWindow) ===" -ForegroundColor Cyan
Write-Host ("  sinceUtc:        {0}" -f $poll.sinceUtc)
Write-Host ("  untilUtc:        {0}" -f $poll.untilUtc)
Write-Host ("  sinceDisplay:    {0} ({1})" -f $poll.sinceDisplay, $poll.displayTimeZoneId)
Write-Host ("  untilDisplay:    {0}" -f $poll.untilDisplay)
Write-Host ("  watermarkBefore: {0}" -f $(if ($poll.watermarkBefore) { $poll.watermarkBefore } else { '(none)' }))
Write-Host "  (First run: sinceUtc uses initialLookbackSeconds from appsettings when watermark is absent.)" -ForegroundColor DarkGray

# Connect PW
$pw = $config.projectWise
$ds = [string]$pw.datasourceName
$credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
Write-Host "`n=== ProjectWise dms_audt ===" -ForegroundColor Cyan
Write-Host "  Datasource: $ds"

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw $connRes.Message }
Write-Host "  Connected." -ForegroundColor Green

$nameFilter = ''
if (-not [string]::IsNullOrWhiteSpace($ItemNameLike)) {
    $esc = $ItemNameLike.Replace("'", "''")
    $nameFilter = " AND o_itemname LIKE '$esc'"
}

$sql = @"
SELECT TOP $($Top * 3) o_acttime, o_action, o_objtype, o_objguid, o_itemname
FROM dms_audt
WHERE o_acttime >= '$sinceUtc'$nameFilter
ORDER BY o_acttime DESC
"@

try {
    $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
    $rows = @()
    if ($result.PSObject.Properties.Name -contains 'Rows' -and $result.Rows) {
        $rows = @($result.Rows)
    } elseif ($result -is [System.Data.DataTable]) {
        $rows = @($result.Rows)
    } else {
        $rows = @($result)
    }
    Write-Host ("  Rows returned: {0}" -f $rows.Count)

    $actionNames = @{
        1020 = 'DOCUMENT_DELETE'; 1007 = 'DOCUMENT_CIN'; 1006 = 'DOCUMENT_FILE_REP'
        1002 = 'DOCUMENT_MODIFY'; 1003 = 'DOCUMENT_ATTR'; 1012 = 'DOCUMENT_STATE'
    }

    $n = 0
    foreach ($row in $rows) {
        if ($n -ge $Top) { break }
        $n++
        $actRaw = $null
        $action = 0
        $item = ''
        $guid = ''
        foreach ($col in $row.Table.Columns) {
            $cn = [string]$col.ColumnName
            $v = $row[$col]
            switch -Regex ($cn) {
                '^o_acttime$' { $actRaw = $v }
                '^o_action$'  { try { $action = [int]$v } catch { } }
                '^o_itemname$'  { $item = [string]$v }
                '^o_objguid$'   { $guid = [string]$v }
            }
        }
        $actName = if ($actionNames.ContainsKey($action)) { $actionNames[$action] } else { "ACTION_$action" }
        $desc = _DescribeActTime $actRaw

        Write-Host ""
        Write-Host ("--- [{0}] {1}  {2}" -f $n, $actName, $item) -ForegroundColor Yellow
        Write-Host ("    o_objguid: {0}" -f $guid)
        Write-Host ("    PW raw:    type={0}  value={1}" -f $desc.type, $desc.raw)
        if ($desc.kind) {
            Write-Host ("    DateTime Kind: {0}" -f $desc.kind)
            Write-Host ("    UTC (assumed): {0}" -f $desc.utcAssumed)
            Write-Host ("    Display zone:  {0}" -f $desc.display)
            Write-Host ("    ToString:      {0}" -f $desc.toStringLocal)
            if ($desc.wrongLocalAsUtc) {
                Write-Host ("    (misread):     {0}" -f $desc.wrongLocalAsUtc) -ForegroundColor DarkGray
            }
        }
        $stored = _FmtPipelinePwActTime -Raw $actRaw
        $inPoll = $false
        if ($actRaw -is [DateTime] -and $poll.untilUtc -and $poll.sinceUtc) {
            $pwStr = (_PwActTimeToUtc $actRaw).ToString('yyyy-MM-dd HH:mm:ss')
            $inPoll = ($pwStr -gt $poll.sinceUtc) -and ($pwStr -le $poll.untilUtc)
        }
        Write-Host ("    audit_events pw_acttime (pipeline): {0}" -f $stored) -ForegroundColor Green
        Write-Host ("    in current poll window (UTC):       {0}" -f $(if ($inPoll) { 'YES' } else { 'no' }))
    }

    if ($n -eq 0) {
        Write-Host "  No rows in window. Widen -Hours or check ItemNameLike." -ForegroundColor DarkYellow
    }
} finally {
    try { Disconnect-PW | Out-Null } catch { }
}

Write-Host "`nDone. PW o_acttime uses UTC wall clock; poll bounds use UTC; pw_acttime/logs use runtime.displayTimeZoneId ($($displayTz.Id))." -ForegroundColor Cyan
