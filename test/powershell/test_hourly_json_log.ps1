$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force

$tmpdir = Join-Path $env:TEMP ('qc_log_test_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpdir -Force | Out-Null
try {
    $env:QC_JSON_LOG_DIR = $tmpdir
    $env:QC_JSON_LOG_TAG = 'TestWatcher'
    Write-QCJsonLog -Level Information -Code 'TEST_ONE' -Message 'first' -Data @{ n = 1 }
    Write-QCJsonLog -Level Information -Code 'TEST_TWO' -Message 'second' -Data @{ n = 2 }
    $hour = Get-QCLogHourStamp
    $f = Join-Path $tmpdir ("TestWatcher_${hour}.jsonl")
    if (-not (Test-Path -LiteralPath $f)) { throw "missing $f" }
    $lines = @(Get-Content -LiteralPath $f)
    if ($lines.Count -ne 2) { throw "expected 2 lines got $($lines.Count)" }
    Write-Host 'test_hourly_json_log: PASS' -ForegroundColor Green
} finally {
    Close-QCJsonLogWriter
    Remove-Item -Recurse -Force $tmpdir -ErrorAction SilentlyContinue
    Remove-Item Env:QC_JSON_LOG_DIR -ErrorAction SilentlyContinue
    Remove-Item Env:QC_JSON_LOG_TAG -ErrorAction SilentlyContinue
}
