# Unit-style checks for audit capture watermark / poll window (no PW connection).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("qc-audit-wm-" + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$wmPath = Join-Path $tmp 'audit-capture-watermark.txt'
$config = @{ database = @{ enabled = $false } }

try {
    $configInitial = @{
        database = @{ enabled = $false }
        auditPoller = @{ initialLookbackSeconds = 3600 }
    }
    $w0 = Get-AuditTrailPollWindow -Config $configInitial -WatermarkPath $wmPath -LookbackSeconds 60
    _Assert ($w0.isFirstCapture) 'first run with initialLookback config'
    _Assert ([int]$w0.lookbackSecondsUsed -eq 3600) 'first run should use initialLookbackSeconds not steady lookback'
    _Assert (($w0.until - $w0.since).TotalSeconds -ge 3599) 'first run should span initial lookback window'

    $w1 = Get-AuditTrailPollWindow -Config $config -WatermarkPath $wmPath -LookbackSeconds 60
    _Assert ($w1.isFirstCapture) 'first run should have no prior capture'
    _Assert (($w1.until - $w1.since).TotalSeconds -ge 59) 'first run should use lookback window'

    $captured = Get-Date '2026-05-26 10:00:00'
    _Assert (Set-AuditTrailCaptureWatermark -WatermarkPath $wmPath -CapturedThrough $captured) 'watermark write should succeed'

    $w2 = Get-AuditTrailPollWindow -Config $config -WatermarkPath $wmPath -LookbackSeconds 60
    _Assert (-not $w2.isFirstCapture) 'second run should see prior capture'
    _Assert ($w2.since -eq $captured.AddSeconds(-60)) 'since should overlap last capture by lookbackSeconds'
    _Assert ($w2.watermarkBefore -eq '2026-05-26 10:00:00') 'watermarkBefore string should match'

    $read = Get-AuditTrailCaptureWatermark -Config $config -WatermarkPath $wmPath
    _Assert ($read -eq $captured) 'read watermark should match written value'

    # Missing watermark file => no capture point (DB watermark ignored when file absent).
    Remove-Item -LiteralPath $wmPath -Force
    $readMissing = Get-AuditTrailCaptureWatermark -Config $config -WatermarkPath $wmPath
    _Assert ($null -eq $readMissing) 'missing watermark file should not use poll_runs-only state'
    $w3 = Get-AuditTrailPollWindow -Config $configInitial -WatermarkPath $wmPath -LookbackSeconds 60
    _Assert ($w3.isFirstCapture) 'no watermark file should trigger initial lookback'

    Write-Host 'OK: audit poll window / watermark tests passed.' -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
