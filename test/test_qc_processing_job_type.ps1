# Maps queue job types to processing_jobs.job_type for telemetry.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/Core.Database.psm1" -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

Assert-Eq (Get-QCProcessingJobType -QueueJobType 'QC_COMMENT_STATUS_SYNC') 'QC_STATE' 'Comment sync maps to QC_STATE'
Assert-Eq (Get-QCProcessingJobType -QueueJobType 'QC_PREPEND') 'QC_PREPEND' 'Other types pass through'

Write-Host 'OK test_qc_processing_job_type.ps1'
