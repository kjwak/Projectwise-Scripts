# Unit tests for QC_COMMENT_STATUS_SYNC trigger rule matching.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/Core.Results.psm1" -Force
Import-Module "$PSScriptRoot/../modules/Core.Paths.psm1" -Force
Import-Module "$PSScriptRoot/../modules/QC.Triggers.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$rule = @{
    id = 'qc-comment-status-pw'
    jobType = 'QC_COMMENT_STATUS_SYNC'
    triggerType = 'pw'
    when = @{
        extensions = @('.pdf')
        fileNameRegexAny = @('(?i)-qc\.pdf$')
    }
    requireAll = @('extensions', 'fileNameRegexAny')
    exclude = @{ pathRegexAny = @(); fileNameRegexAny = @() }
}

$matchCandidate = @{
    path = 'Documents\Proj\CADD\Sheets\A101-qc.pdf'
    fileName = 'A101-qc.pdf'
    description = ''
}

$noMatchCandidate = @{
    path = 'Documents\Proj\CADD\Sheets\A101.pdf'
    fileName = 'A101.pdf'
    description = '|QC|'
}

$r1 = Test-TriggerRule -Candidate $matchCandidate -Rule $rule
Assert-True $r1.IsSuccess 'Trigger test should succeed'
Assert-True ([bool]$r1.Data.isMatch) 'A101-qc.pdf should match -qc.pdf rule'

$r2 = Test-TriggerRule -Candidate $noMatchCandidate -Rule $rule
Assert-True $r2.IsSuccess 'Trigger test should succeed for non-match'
Assert-True (-not [bool]$r2.Data.isMatch) 'A101.pdf should not match -qc.pdf rule'

Write-Host 'OK test_qc_comment_trigger.ps1'
