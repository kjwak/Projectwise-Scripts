<#
.SYNOPSIS
Regression: negative stamp coordinates must be quoted on the overlay CLI command line.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.ReviewStamp.psm1') -Force -DisableNameChecking | Out-Null

function Assert-True($Cond, $Msg) {
    if (-not $Cond) { throw "ASSERT FAILED: $Msg" }
}

$line = Join-QCProcessArgumentList -ArgumentTokens @(
    '--apply-review-stamp', 'C:\out\merged.pdf', 'C:\stamps\Peer.pdf',
    '--stamp-x-pt', '-400', '--stamp-y-pt', '0', '--stamp-height-pt', '300'
)

Assert-True ($line -match '--stamp-x-pt\s+"-400"') "Expected quoted -400 after --stamp-x-pt; got: $line"
Assert-True ($line -match '--stamp-y-pt\s+0\b') "Expected unquoted 0 for --stamp-y-pt; got: $line"

Write-Host 'test_review_stamp_negative_cli: PASS' -ForegroundColor Green
