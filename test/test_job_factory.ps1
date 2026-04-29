# Unit tests for QC.JobFactory (pure logic).
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

$config = @{
    jobFactory = @{
        idPolicy = 'deterministic'
        createdAtUtc = '2026-01-01T00:00:00.0000000Z'
    }
}

$candidate = @{
    path = 'Documents\AZDOT 2024\ProjA\CADD\Sheets\A101.pdf'
    fileName = 'A101.pdf'
    description = '|QC| ready'
}

$rule = @{
    id = 'rule-1'
    jobType = 'QC_PREPEND'
    triggerType = 'desc'
}

# deterministic job id
$id1 = New-QCJobId -Candidate $candidate -Rule $rule -Config $config
Assert-True $id1.IsSuccess 'New-QCJobId should succeed'
$id2 = New-QCJobId -Candidate $candidate -Rule $rule -Config $config
Assert-True $id2.IsSuccess 'New-QCJobId should succeed (second call)'
Assert-Eq $id1.Data.id $id2.Data.id 'JobId should be deterministic for same inputs'

# deterministic dedupe key
$jobSeed = @{
    type = 'QC_PREPEND'
    sourcePath = $candidate.path
    triggerRule = @{ id = $rule.id }
}
$d1 = Get-QCDedupeKey -Job $jobSeed -Config $config
Assert-True $d1.IsSuccess 'Get-QCDedupeKey should succeed'
$d2 = Get-QCDedupeKey -Job $jobSeed -Config $config
Assert-True $d2.IsSuccess 'Get-QCDedupeKey should succeed (second call)'
Assert-Eq $d1.Data.dedupeKey $d2.Data.dedupeKey 'DedupeKey should be deterministic for same inputs'

# required field validation: clean failure on missing type
$badJob = @{
    sourcePath = 'Documents\X\Y.pdf'
    triggerRule = @{ id = 'r' }
    dedupeKey = 'dq_x'
}
$v = Test-QCJobRequiredFields -Job $badJob
Assert-True (-not $v.IsSuccess) 'Missing type should fail validation'
Assert-Eq $v.Code 'JOB_VALIDATION_MISSING_REQUIRED_FIELDS' 'Missing required fields should return expected code'
Assert-True ($v.Data.missing -contains 'type') 'Missing list should include type'

# required field validation: clean failure on missing sourcePath
$badJob2 = @{
    type = 'QC_PREPEND'
    triggerRule = @{ id = 'r' }
    dedupeKey = 'dq_x'
}
$v2 = Test-QCJobRequiredFields -Job $badJob2
Assert-True (-not $v2.IsSuccess) 'Missing sourcePath should fail validation'
Assert-True ($v2.Data.missing -contains 'sourcePath') 'Missing list should include sourcePath'

# required field validation: clean failure on missing triggerRule
$badJob3 = @{
    type = 'QC_PREPEND'
    sourcePath = 'Documents\X\Y.pdf'
    dedupeKey = 'dq_x'
}
$v3 = Test-QCJobRequiredFields -Job $badJob3
Assert-True (-not $v3.IsSuccess) 'Missing triggerRule should fail validation'
Assert-True ($v3.Data.missing -contains 'triggerRule') 'Missing list should include triggerRule'

# valid job object creation
$jr = New-QCJobObject -Candidate $candidate -Rule $rule -Config $config
Assert-True $jr.IsSuccess 'New-QCJobObject should succeed'
$job = $jr.Data.job
foreach ($k in @('id','type','sourcePath','sourceName','triggerRule','dedupeKey','status','createdAt','attempts','metadata')) {
    Assert-True ($job.ContainsKey($k)) "Job should contain key: $k"
}
Assert-Eq $job.type 'QC_PREPEND' 'Job.type should match rule.jobType'
Assert-Eq $job.triggerRule.id 'rule-1' 'Job.triggerRule.id should match rule.id'
Assert-Eq $job.status 'queued' 'Job.status should be queued'
Assert-Eq $job.attempts 0 'Job.attempts should start at 0'
Assert-True (-not [string]::IsNullOrWhiteSpace([string]$job.dedupeKey)) 'Job should have dedupeKey'

Write-Host 'All QC.JobFactory tests passed.' -ForegroundColor Green

