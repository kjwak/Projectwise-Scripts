<#
.SYNOPSIS
Validates SQL Server connectivity, schema initialization, and basic CRUD operations.

.DESCRIPTION
End-to-end test of the QC pipeline database setup:
  1. Loads appsettings.json
  2. Tests raw SqlClient connectivity
  3. Runs Initialize-QCDatabaseSchema (idempotent)
  4. Verifies all expected tables exist
  5. Verifies all expected views exist
  6. Tests INSERT/SELECT/DELETE round-trip
  7. Reports schema_version

.PARAMETER AppSettingsPath
Path to appsettings.json. Defaults to repo root.

.PARAMETER Pretty
Emit formatted JSON output.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [switch]$Pretty
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

# Import modules with explicit error trapping
$modulesDir = Join-Path $repoRoot 'modules'
foreach ($mod in @('Core.Results.psm1', 'Core.Database.psm1')) {
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

$warnings = [System.Collections.Generic.List[string]]::new()
$checks = [ordered]@{}

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
$checks['configLoaded'] = $true
$checks['databaseEnabled'] = $dbEnabled
Write-Host "  database.enabled = $dbEnabled" -ForegroundColor $(if ($dbEnabled) { 'Green' } else { 'Red' })

if (-not $dbEnabled) {
    Write-Host "  BLOCKED: database is not enabled in config." -ForegroundColor Red
    $warnings.Add('database.enabled is false in appsettings.json')
}

# -- 1. Raw connectivity test --------------------------------------------------

Write-Host "`n[1] Testing SQL Server connectivity..." -ForegroundColor Yellow
$connRes = Get-QCDatabaseConnection -Config $config
if ($connRes.IsSuccess) {
    $conn = $connRes.Data.connection
    $serverVersion = $conn.ServerVersion
    $database = $conn.Database
    $conn.Close()
    $conn.Dispose()
    $checks['connectivity'] = $true
    $checks['serverVersion'] = $serverVersion
    $checks['database'] = $database
    Write-Host "  Connected to $database (SQL Server $serverVersion)" -ForegroundColor Green
} else {
    $checks['connectivity'] = $false
    $checks['connectError'] = $connRes.Message
    $warnings.Add("Connection failed: $($connRes.Message)")
    Write-Host "  FAILED: $($connRes.Message)" -ForegroundColor Red
    Write-Host "`nCannot proceed without a database connection." -ForegroundColor Red
    $result = [ordered]@{ timestamp = (Get-Date).ToString('o'); checks = $checks; warnings = @($warnings) }
    if ($Pretty) { $result | ConvertTo-Json -Depth 10 } else { $result | ConvertTo-Json -Depth 10 -Compress }
    return
}

# -- 2. Schema initialization -------------------------------------------------

Write-Host "`n[2] Running Initialize-QCDatabaseSchema..." -ForegroundColor Yellow
$schemaRes = Initialize-QCDatabaseSchema -Config $config
$checks['schemaInit'] = $schemaRes.IsSuccess
$checks['schemaCode'] = $schemaRes.Code
$checks['schemaMessage'] = $schemaRes.Message
if ($schemaRes.IsSuccess) {
    Write-Host "  $($schemaRes.Message)" -ForegroundColor Green
} else {
    $warnings.Add("Schema init failed: $($schemaRes.Message)")
    Write-Host "  FAILED: $($schemaRes.Message)" -ForegroundColor Red
}

# -- 3. Verify tables exist ---------------------------------------------------

Write-Host "`n[3] Verifying tables..." -ForegroundColor Yellow
$expectedTables = @(
    'schema_version',
    'audit_events',
    'document_activity',
    'document_state_history',
    'transition_events',
    'poll_runs',
    'processing_jobs',
    'notification_log'
)

$tableRes = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
$foundTables = @()
if ($tableRes.IsSuccess) {
    $foundTables = @($tableRes.Data.table.Rows | ForEach-Object { [string]$_.TABLE_NAME })
}

$tableMissing = [System.Collections.Generic.List[string]]::new()
foreach ($t in $expectedTables) {
    $exists = $foundTables -contains $t
    if ($exists) {
        Write-Host "  [OK ] $t" -ForegroundColor Green
    } else {
        Write-Host "  [NO ] $t" -ForegroundColor Red
        $tableMissing.Add($t)
    }
}
$checks['tablesExpected'] = $expectedTables.Count
$checks['tablesFound'] = $foundTables.Count
$checks['tablesMissing'] = @($tableMissing)

# -- 4. Verify views exist ----------------------------------------------------

Write-Host "`n[4] Verifying views..." -ForegroundColor Yellow
$expectedViews = @(
    'v_qc_cycle_aging',
    'v_folder_activity',
    'v_poller_health',
    'v_job_summary'
)

$viewRes = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS ORDER BY TABLE_NAME"
$foundViews = @()
if ($viewRes.IsSuccess) {
    $foundViews = @($viewRes.Data.table.Rows | ForEach-Object { [string]$_.TABLE_NAME })
}

$viewsMissing = [System.Collections.Generic.List[string]]::new()
foreach ($v in $expectedViews) {
    $exists = $foundViews -contains $v
    if ($exists) {
        Write-Host "  [OK ] $v" -ForegroundColor Green
    } else {
        Write-Host "  [NO ] $v" -ForegroundColor Red
        $viewsMissing.Add($v)
    }
}
$checks['viewsExpected'] = $expectedViews.Count
$checks['viewsFound'] = $foundViews.Count
$checks['viewsMissing'] = @($viewsMissing)

# -- 5. CRUD round-trip test ---------------------------------------------------

Write-Host "`n[5] CRUD round-trip test (poll_runs)..." -ForegroundColor Yellow
$testTs = (Get-Date).ToString('o')
$insertRes = Invoke-QCDatabaseNonQuery -Config $config -Sql "INSERT INTO poll_runs (started_at, events_fetched, events_relevant, candidates_created, jobs_enqueued, error_message, is_reconciliation) VALUES (@ts, 0, 0, 0, 0, 'connectivity_test', 0)" -Parameters @{ ts = $testTs }

if ($insertRes.IsSuccess) {
    Write-Host "  INSERT: OK" -ForegroundColor Green
    $selectRes = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT TOP 1 id, started_at, error_message FROM poll_runs WHERE error_message = 'connectivity_test' ORDER BY id DESC"
    if ($selectRes.IsSuccess -and $selectRes.Data.rowCount -gt 0) {
        $row = $selectRes.Data.table.Rows[0]
        Write-Host "  SELECT: OK (id=$($row.id))" -ForegroundColor Green
        $deleteRes = Invoke-QCDatabaseNonQuery -Config $config -Sql "DELETE FROM poll_runs WHERE error_message = 'connectivity_test'"
        if ($deleteRes.IsSuccess) {
            Write-Host "  DELETE: OK ($($deleteRes.Data.rowsAffected) rows cleaned up)" -ForegroundColor Green
            $checks['crudTest'] = 'PASS'
        } else {
            $checks['crudTest'] = 'DELETE_FAILED'
            $warnings.Add("CRUD delete failed: $($deleteRes.Message)")
        }
    } else {
        $checks['crudTest'] = 'SELECT_FAILED'
        $warnings.Add("CRUD select failed")
    }
} else {
    $checks['crudTest'] = 'INSERT_FAILED'
    $warnings.Add("CRUD insert failed: $($insertRes.Message)")
    Write-Host "  FAILED: $($insertRes.Message)" -ForegroundColor Red
}

# -- 6. Schema version --------------------------------------------------------

Write-Host "`n[6] Schema version..." -ForegroundColor Yellow
$verRes = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT version, applied_at, description FROM schema_version ORDER BY id DESC"
if ($verRes.IsSuccess -and $verRes.Data.rowCount -gt 0) {
    $latest = $verRes.Data.table.Rows[0]
    $checks['schemaVersion'] = [string]$latest.version
    $checks['schemaAppliedAt'] = [string]$latest.applied_at
    Write-Host "  Version: $($latest.version) (applied: $($latest.applied_at))" -ForegroundColor Green
} else {
    $checks['schemaVersion'] = $null
    Write-Host "  No schema version records found." -ForegroundColor DarkYellow
}

# -- Summary -------------------------------------------------------------------

$allOk = $checks['connectivity'] -and $checks['schemaInit'] -and ($tableMissing.Count -eq 0) -and ($viewsMissing.Count -eq 0) -and ($checks['crudTest'] -eq 'PASS')

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Connectivity:    $(if ($checks['connectivity']) { 'OK' } else { 'FAILED' })"
Write-Host "Schema init:     $($checks['schemaCode'])"
Write-Host "Tables:          $($checks['tablesFound'])/$($checks['tablesExpected'])"
Write-Host "Views:           $($checks['viewsFound'])/$($checks['viewsExpected'])"
Write-Host "CRUD test:       $($checks['crudTest'])"
Write-Host "Schema version:  $($checks['schemaVersion'])"
Write-Host "Warnings:        $($warnings.Count)"
Write-Host "Overall:         $(if ($allOk) { 'PASS' } else { 'ISSUES DETECTED' })" -ForegroundColor $(if ($allOk) { 'Green' } else { 'Red' })

if ($warnings.Count -gt 0) {
    Write-Host "`nWarnings:" -ForegroundColor DarkYellow
    foreach ($w in $warnings) { Write-Host "  - $w" -ForegroundColor DarkYellow }
}

$result = [ordered]@{
    timestamp       = (Get-Date).ToString('o')
    overall         = if ($allOk) { 'PASS' } else { 'ISSUES' }
    checks          = $checks
    warnings        = @($warnings)
}

$depth = 10
if ($Pretty) { $result | ConvertTo-Json -Depth $depth }
else { $result | ConvertTo-Json -Depth $depth -Compress }
