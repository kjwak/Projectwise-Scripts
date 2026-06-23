# Final prepend stamp gating: check/review stamped; production skipped.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.ReviewStamp.psm1') -Force

function Assert-True($c, $msg) { if (-not $c) { throw "ASSERT FAILED: $msg" } }
function Assert-False($c, $msg) { if ($c) { throw "ASSERT FAILED: $msg" } }

Assert-False (Test-QCPrependSkipReviewStamp -PrependTrigger 'initialQcPdf' -ProcessType 'check') `
    'initial check prepend should not skip stamp'
Assert-False (Test-QCPrependSkipReviewStamp -PrependTrigger '' -ProcessType 'review') `
    'empty trigger should not skip stamp'

Assert-False (Test-QCPrependSkipReviewStamp -PrependTrigger 'finalQcComplete' -ProcessType 'check') `
    'final check prepend should apply stamp'
Assert-False (Test-QCPrependSkipReviewStamp -PrependTrigger 'finalQcComplete' -ProcessType 'review') `
    'final review prepend should apply stamp'
Assert-False (Test-QCPrependSkipReviewStamp -PrependTrigger 'qcFinalizing' -ProcessType 'Check') `
    'final check prepend should apply stamp (normalized case)'

Assert-True (Test-QCPrependSkipReviewStamp -PrependTrigger 'finalQcComplete' -ProcessType 'production') `
    'final production prepend should skip stamp'
Assert-True (Test-QCPrependSkipReviewStamp -PrependTrigger 'finalPrepend' -ProcessType 'prod') `
    'final production prepend should skip stamp (prod alias)'
Assert-True (Test-QCPrependSkipReviewStamp -PrependTrigger 'finalQcComplete' -ProcessType '') `
    'final prepend with unknown process type should skip stamp'

Write-Host 'All QC prepend stamp skip tests passed.' -ForegroundColor Green
