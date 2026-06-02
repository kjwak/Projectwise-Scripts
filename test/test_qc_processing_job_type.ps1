# Maps queue job types to processing_jobs.job_type for telemetry.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/Core.Database.psm1" -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

Assert-Eq (Get-QCProcessingJobType -QueueJobType 'QC_COMMENT_STATUS_SYNC') 'QC_STATE' 'Comment sync maps to QC_STATE'
Assert-Eq (Get-QCProcessingJobType -QueueJobType 'QC_STATE') 'QC_STATE' 'QC_STATE passes through'
Assert-Eq (Get-QCProcessingJobType -QueueJobType 'QC_PREPEND') 'QC_PREPEND' 'Other types pass through'

$id1 = New-QCStateChangeJobId -ParentJobId 'job-abc' -DocumentGuid 'g1' -PreviousState 'In Production' -CurrentState 'QC Received'
Assert-Eq $id1 'job-abc|state' 'Parent job id suffix for state telemetry'

$id2 = New-QCStateChangeJobId -DocumentGuid '7e635293-0c08-4d88-b183-c4634be7763' -PreviousState 'A' -CurrentState 'B'
if ($id2 -notlike 'qc-state-*') { throw "Expected qc-state-* id, got $id2" }

Write-Host 'OK test_qc_processing_job_type.ps1'
