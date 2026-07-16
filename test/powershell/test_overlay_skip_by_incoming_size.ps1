# Unit tests for Test-QCOverlaySkipByIncomingSize (overlay size gate).
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.PrependOverlayPolicy.psm1') -Force

$mb = 1024L * 1024L

Write-Host 'Test: threshold disabled (missing key)' -ForegroundColor Cyan
$r = Test-QCOverlaySkipByIncomingSize -QcPrepend @{} -IncomingBytes (100L * $mb)
Assert-Eq $r.Skip $false 'missing key should not skip'
Assert-Eq $r.Reason 'threshold_disabled' 'reason for missing key'

Write-Host 'Test: threshold disabled (0)' -ForegroundColor Cyan
$r = Test-QCOverlaySkipByIncomingSize -QcPrepend @{ overlaySkipIncomingAboveMb = 0 } -IncomingBytes (100L * $mb)
Assert-Eq $r.Skip $false '0 should not skip'
Assert-Eq $r.Reason 'threshold_disabled' 'reason for 0'

Write-Host 'Test: within threshold (exact 5 MB is not skipped)' -ForegroundColor Cyan
$r = Test-QCOverlaySkipByIncomingSize -QcPrepend @{ overlaySkipIncomingAboveMb = 5 } -IncomingBytes (5L * $mb)
Assert-Eq $r.Skip $false 'exact threshold should keep overlay'
Assert-Eq $r.Reason 'incoming_within_threshold' 'within reason'
Assert-Eq $r.ThresholdBytes (5L * $mb) 'threshold bytes'

Write-Host 'Test: above threshold skips' -ForegroundColor Cyan
$r = Test-QCOverlaySkipByIncomingSize -QcPrepend @{ overlaySkipIncomingAboveMb = 5 } -IncomingBytes ((5L * $mb) + 1L)
Assert-Eq $r.Skip $true 'one byte over should skip'
Assert-Eq $r.Reason 'incoming_above_threshold' 'above reason'
Assert-Eq $r.IncomingBytes ((5L * $mb) + 1L) 'incoming bytes echoed'

Write-Host 'Test: small file under default-like 5 MB' -ForegroundColor Cyan
$r = Test-QCOverlaySkipByIncomingSize -QcPrepend @{ overlaySkipIncomingAboveMb = 5 } -IncomingBytes (512L * 1024L)
Assert-Eq $r.Skip $false '512 KiB should keep overlay'

Write-Host 'Test: fractional MB threshold' -ForegroundColor Cyan
$r = Test-QCOverlaySkipByIncomingSize -QcPrepend @{ overlaySkipIncomingAboveMb = 0.5 } -IncomingBytes (600L * 1024L)
Assert-Eq $r.Skip $true '600 KiB > 0.5 MB should skip'

Write-Host 'ALL PASSED' -ForegroundColor Green
