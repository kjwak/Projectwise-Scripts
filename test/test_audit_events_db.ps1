<#
.SYNOPSIS
Integration test for audit_events ingestion (Write-QCAuditEventRows).

.DESCRIPTION
Requires appsettings database.enabled=true and a reachable SQL Server.
Skips with exit 0 when database is disabled (CI-friendly).

.EXAMPLE
.\test\test_audit_events_db.ps1
.\test\test_audit_events_db.ps1 -AppSettingsPath .\appsettings.json
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = ''
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not (Test-QCDatabaseEnabled -Config $config)) {
    Write-Host 'SKIP: database.enabled is false in appsettings.' -ForegroundColor Yellow
    exit 0
}

$connProbe = Get-QCDatabaseConnection -Config $config
if (-not $connProbe.IsSuccess) {
    Write-Host "SKIP: cannot connect to SQL Server ($($connProbe.Message))." -ForegroundColor Yellow
    exit 0
}
try { $connProbe.Data.connection.Close(); $connProbe.Data.connection.Dispose() } catch { }

$schemaRes = Initialize-QCDatabaseSchema -Config $config
if (-not $schemaRes.IsSuccess) { throw "Schema init failed: $($schemaRes.Message)" }

$testGuid = 'TEST-AUDIT-' + [guid]::NewGuid().ToString('N')
$actTime = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
$marker = "QC_TEST_$testGuid"

$rows = @(
    @{
        acttime       = $actTime
        action        = 1002
        actionName    = 'DOCUMENT_MODIFY'
        objtype       = 2
        objno         = 0
        objguid       = $testGuid
        parentguid    = 'TEST-PARENT-00000000-0000-0000-0000-000000000001'
        userno        = 0
        itemname      = 'test-sheet.pdf'
        itemdesc      = $marker
        textparam     = $null
        folder        = 'Documents\QC_TEST\Proj\CADD\Sheets'
        candidateType = 'WATCH_MATCH'
    }
)

try {
    $w1 = Write-QCAuditEventRows -Config $config -Rows $rows
    _Assert $w1.IsSuccess "first insert should succeed: $($w1.Message)"
    _Assert ([int]$w1.Data.written -eq 1) "first insert should write 1 row (got $($w1.Data.written))"

    $w2 = Write-QCAuditEventRows -Config $config -Rows $rows
    _Assert $w2.IsSuccess "duplicate insert should succeed (dedupe): $($w2.Message)"
    _Assert ([int]$w2.Data.written -eq 0) "duplicate insert should write 0 rows (got $($w2.Data.written))"
    _Assert ([int]$w2.Data.skipped -ge 1) 'duplicate should be counted as skipped'

    $q = Invoke-QCDatabaseScalar -Config $config -Sql @"
SELECT COUNT(*) FROM audit_events
WHERE pw_objguid = @g AND pw_itemdesc = @m
"@ -Parameters @{ g = $testGuid; m = $marker }
    _Assert $q.IsSuccess "count query failed: $($q.Message)"
    _Assert ([int]$q.Data.value -eq 1) "expected exactly 1 row for test guid (got $($q.Data.value))"

    # Null objguid rows must not insert (production filter).
    $bad = Write-QCAuditEventRows -Config $config -Rows @(@{
            acttime = $actTime; action = 1002; actionName = 'DOCUMENT_MODIFY'; objtype = 2
            objno = 0; objguid = $null; parentguid = $null; userno = 0
            itemname = 'x.pdf'; itemdesc = $marker; textparam = $null; folder = $null; candidateType = $null
        })
    _Assert $bad.IsSuccess 'null guid batch should not throw'
    _Assert ([int]$bad.Data.written -eq 0) 'null objguid row must not insert'

    Write-Host 'OK: audit_events DB insert + dedupe tests passed.' -ForegroundColor Green
    exit 0
} finally {
    try {
        $null = Invoke-QCDatabaseNonQuery -Config $config -Sql 'DELETE FROM audit_events WHERE pw_objguid = @g' -Parameters @{ g = $testGuid }
    } catch { }
}
