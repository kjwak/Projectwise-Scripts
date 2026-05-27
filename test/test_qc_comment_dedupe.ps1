# Unit tests for QC_COMMENT_STATUS_SYNC dedupe keys.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/Core.Results.psm1" -Force
Import-Module "$PSScriptRoot/../modules/Core.Paths.psm1" -Force
Import-Module "$PSScriptRoot/../modules/QC.JobFactory.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$config = @{ jobFactory = @{ idPolicy = 'deterministic'; createdAtUtc = '2026-01-01T00:00:00.0000000Z' } }
$rule = @{
    id = 'qc-comment-status-pw'
    jobType = 'QC_COMMENT_STATUS_SYNC'
    triggerType = 'pw'
    grouping = @{ enabled = $false; groupBy = 'file' }
}

$jobSeed = @{
    type = 'QC_COMMENT_STATUS_SYNC'
    sourcePath = 'Documents\Proj\Sheets\A101-qc.pdf'
    triggerRule = @{ id = $rule.id; grouping = $rule.grouping }
    metadata = @{ candidate = @{ file = @{ sha256 = 'hash-aaa' } } }
}

$d1 = Get-QCDedupeKey -Job $jobSeed -Config $config
$d2 = Get-QCDedupeKey -Job $jobSeed -Config $config
Assert-True $d1.IsSuccess 'Dedupe should succeed'
Assert-Eq $d1.Data.dedupeKey $d2.Data.dedupeKey 'Dedupe key should be stable'

$jobSeed2 = @{
    type = 'QC_COMMENT_STATUS_SYNC'
    sourcePath = 'Documents\Proj\Sheets\A101-qc.pdf'
    triggerRule = @{ id = $rule.id; grouping = $rule.grouping }
    metadata = @{ candidate = @{ file = @{ sha256 = 'hash-bbb' } } }
}
$d3 = Get-QCDedupeKey -Job $jobSeed2 -Config $config
Assert-True ($d1.Data.dedupeKey -ne $d3.Data.dedupeKey) 'Different file hash should change dedupe key'

Write-Host 'OK test_qc_comment_dedupe.ps1'
