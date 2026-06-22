# Unit-style checks for audit capture watermark / poll window (no PW connection).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.AuditPoller.psm1') -Force

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

    $captured = [DateTime]::SpecifyKind([DateTime]::Parse('2026-05-26 10:00:00'), [DateTimeKind]::Utc)
    _Assert (Set-AuditTrailCaptureWatermark -WatermarkPath $wmPath -CapturedThrough $captured) 'watermark write should succeed'
    $wmRaw = (Get-Content -LiteralPath $wmPath -Raw).Trim()
    _Assert ($wmRaw.EndsWith('Z')) 'watermark file should store UTC with Z suffix'

    $configOverlap = @{
        database = @{ enabled = $false }
        auditPoller = @{ overlapSeconds = 5 }
    }
    $w2 = Get-AuditTrailPollWindow -Config $configOverlap -WatermarkPath $wmPath -LookbackSeconds 60
    _Assert (-not $w2.isFirstCapture) 'second run should see prior capture'
    _Assert ($w2.since -eq $captured.AddSeconds(-5)) 'since should use overlapSeconds not lookbackSeconds'
    _Assert ([int]$w2.overlapSecondsUsed -eq 5) 'overlapSecondsUsed should reflect config'
    _Assert ($w2.watermarkBefore -eq '2026-05-26 10:00:00') 'watermarkBefore string should match UTC clock'

    $w2b = Get-AuditTrailPollWindow -Config $config -WatermarkPath $wmPath -LookbackSeconds 60
    _Assert ($w2b.since -eq $captured.AddSeconds(-1)) 'default steady-state should rewind 1s at watermark for second-granular dms_audt'
    _Assert ([int]$w2b.overlapSecondsUsed -eq 1) 'implicit 1s overlap when overlapSeconds unset'

    $read = Get-AuditTrailCaptureWatermark -Config $config -WatermarkPath $wmPath
    _Assert ($read.ToUniversalTime() -eq $captured) 'read watermark should match written UTC value'

    Set-Content -LiteralPath $wmPath -Value '2026-06-03 14:10:47' -Encoding UTF8 -NoNewline
    $readNoZ = Get-AuditTrailCaptureWatermark -Config $config -WatermarkPath $wmPath
    _Assert ($readNoZ) 'watermark file without Z suffix should parse'
    _Assert ($readNoZ.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') -eq '2026-06-03 14:10:47') 'non-Z file watermark must stay UTC wall clock (not +7h local)'

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
