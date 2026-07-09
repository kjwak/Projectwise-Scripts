# Unit + optional integration tests for automation_events telemetry.
param(
    [string]$AppSettingsPath = ''
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Telemetry.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT: $msg" } }

$config = @{
    telemetry = @{
        automationEvents = @{
            enabled = $true
            excludeCodes = @('WATCH_TICK_START', 'WATCH_TICK_SLEEP', 'WORKER_NO_JOB')
        }
    }
}

if (Test-QCAutomationEventPersist -Level 'Information' -Code 'WATCH_TICK_START' -Config $config) { throw 'heartbeat should be excluded' }
if (Test-QCAutomationEventPersist -Level 'Information' -Code 'WORKER_NO_JOB' -Config $config) { throw 'WORKER_NO_JOB should be excluded' }
if (-not (Test-QCAutomationEventPersist -Level 'Warning' -Code 'WATCH_TICK_START' -Config $config)) { throw 'warning should always persist' }
if (-not (Test-QCAutomationEventPersist -Level 'Error' -Code 'WORKER_NO_JOB' -Config $config)) { throw 'error should always persist' }
if (Test-QCAutomationEventPersist -Level 'Information' -Code 'WORKER_STAGE' -Message 'polling queue' -Config $config) { throw 'noisy WORKER_STAGE should be excluded' }
if (-not (Test-QCAutomationEventPersist -Level 'Information' -Code 'WORKER_STAGE' -Message 'dispatch' -Data @{ jobId = 'job-1' } -Config $config)) { throw 'WORKER_STAGE with job_id should persist' }

$disabledCfg = @{ telemetry = @{ automationEvents = @{ enabled = $false } } }
if (Test-QCAutomationEventPersist -Level 'Error' -Code 'X' -Config $disabledCfg) { throw 'disabled telemetry should skip' }

# DB disabled must not throw caller
$res = Write-QCAutomationEvent -Level 'Information' -Code 'WORKER_START' -Message 'test' -Config @{ database = @{ enabled = $false } }
Assert-True ($res.IsSuccess) 'disabled db returns success result'
Assert-True ($res.Data.skipped) 'disabled db skips write'

# Regression: Write-QCAutomationEvent must not call Core.Database private helpers
# (_QDB-TruncateTelemetryPayload is module-private and was silently breaking all inserts).
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
$truncateProbe = Write-QCAutomationEvent -Level 'Information' -Code 'WORKER_START' -Message ('truncate-probe-' + [guid]::NewGuid().ToString('n')) -Config @{ database = @{ enabled = $false } }
Assert-True ($truncateProbe.IsSuccess) 'truncate path must not throw CommandNotFound across modules'
Assert-True ($truncateProbe.Code -ne 'AUTOMATION_EVENT_WRITE_FAILED') 'truncate path must not fail when db disabled'

# JSONL import line preserves payload intent
$line = '{"ts":"2026-06-15T10:00:00-07:00","level":"Information","code":"WORKER_SUCCEEDED","message":"ok","data":{"jobId":"j1","documentGuid":"g1"}}'
$parsed = $line | ConvertFrom-Json
Assert-True ($parsed.data.jobId -eq 'j1') 'fixture line parses'

Write-Host 'Unit tests: PASS' -ForegroundColor Green

if (-not (Test-Path -LiteralPath $AppSettingsPath)) { exit 0 }
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Telemetry.psm1') -Force
try {
    $dbConfig = Get-QCAppSettingsConfig -Path $AppSettingsPath
} catch {
    Write-Host 'SKIP integration: could not load appsettings.' -ForegroundColor Yellow
    exit 0
}
if (-not (Test-QCDatabaseEnabled -Config $dbConfig)) {
    Write-Host 'SKIP integration: database disabled.' -ForegroundColor Yellow
    exit 0
}
$connProbe = Get-QCDatabaseConnection -Config $dbConfig
if (-not $connProbe.IsSuccess) {
    Write-Host 'SKIP integration: SQL unavailable.' -ForegroundColor Yellow
    exit 0
}
try { $connProbe.Data.connection.Close(); $connProbe.Data.connection.Dispose() } catch { }

$schemaRes = Initialize-QCDatabaseSchema -Config $dbConfig
if (-not $schemaRes.IsSuccess) { throw $schemaRes.Message }

$marker = 'QC_TELEM_TEST_' + [guid]::NewGuid().ToString('N')
Set-QCAutomationTelemetryContext -Config $dbConfig -ProcessName 'test_automation_events_telemetry' -RunId $marker | Out-Null

try {
    $w1 = Write-QCAutomationEvent -Level 'Information' -Code 'WORKER_START' -Message $marker -Data @{ jobId = $marker; documentGuid = $marker }
    Assert-True ($w1.Data.written) 'first insert written'

    $w2 = Write-QCAutomationEvent -Level 'Information' -Code 'WORKER_START' -Message $marker -Data @{ jobId = $marker; documentGuid = $marker }
    Assert-True ($w2.Code -eq 'AUTOMATION_EVENT_DUPLICATE' -or ($w2.Data.duplicate -eq $true)) 'duplicate insert idempotent'

    $q = Invoke-QCDatabaseScalar -Config $dbConfig -Sql @"
SELECT COUNT(*) FROM automation_events WHERE message = @m AND job_id = @j
"@ -Parameters @{ m = $marker; j = $marker }
    Assert-True ([int]$q.Data.value -eq 1) 'exactly one row after duplicate'

    $viewRows = Invoke-QCDatabaseQuery -Config $dbConfig -Sql "SELECT TOP (5) code, level FROM v_mcp_recent_errors WHERE message = @m" -Parameters @{ m = $marker }
    Assert-True ($viewRows.IsSuccess) 'view query ok'

  # import idempotency from temp jsonl
    $tmp = Join-Path $env:TEMP ('qc_import_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $jsonl = Join-Path $tmp 'Run-QCProcessor_2026-06-15_10.jsonl'
    $importLine = (@{
        ts = (Get-Date).ToString('o')
        level = 'Information'
        code = 'JOB_TELEMETRY_WRITTEN'
        message = $marker
        data = @{ jobId = $marker; runId = $marker }
    } | ConvertTo-Json -Compress)
    Set-Content -LiteralPath $jsonl -Value $importLine -Encoding UTF8
    $i1 = Import-QCAutomationEventFromJsonLine -Line $importLine -ProcessName 'Run-QCProcessor' -Config $dbConfig
    $i2 = Import-QCAutomationEventFromJsonLine -Line $importLine -ProcessName 'Run-QCProcessor' -Config $dbConfig
    Assert-True ($i1.Data.inserted -or $i1.Data.detail.Data.written) 'import first pass'
    Assert-True ($i2.Data.skipped -or $i2.Data.detail.Code -eq 'AUTOMATION_EVENT_DUPLICATE' -or -not $i2.Data.inserted) 'import second pass idempotent'
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue

    Write-Host 'Integration tests: PASS' -ForegroundColor Green
} finally {
    try {
        $null = Invoke-QCDatabaseNonQuery -Config $dbConfig -Sql 'DELETE FROM automation_events WHERE job_id = @j OR message = @m' -Parameters @{ j = $marker; m = $marker }
    } catch { }
}
