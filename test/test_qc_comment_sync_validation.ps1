# Regression tests from comment-sync validation review.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Database/Core.Database.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$mapped = Get-QCProcessingJobType -QueueJobType 'QC_COMMENT_STATUS_SYNC'
Assert-True ($mapped -eq 'QC_STATE') 'Telemetry maps comment sync to QC_STATE only in processing_jobs'

# Dry-run DB writes blocked by default
$cfgDry = @{ dryRun = $true; database = @{ enabled = $true; allowWritesInDryRun = $false } }
Assert-True (-not (Test-QCDatabaseWritesAllowed -Config $cfgDry)) 'Dry-run blocks DB writes by default'
$cfgDryAllow = @{ dryRun = $true; database = @{ enabled = $true; allowWritesInDryRun = $true } }
Assert-True (Test-QCDatabaseWritesAllowed -Config $cfgDryAllow) 'allowWritesInDryRun permits dry-run DB writes'

Import-Module "$repoRoot/modules/Processing/QC.CommentSync.Database.psm1" -Force

# Planned workflow events stay in-memory unless logPlannedEventsInDryRun
$cfgPlanned = @{ dryRun = $true; database = @{ enabled = $true; logPlannedEventsInDryRun = $false } }
$ev = Write-QCWorkflowEvent -Config $cfgPlanned -EventType 'STATE_DECIDED' -PlannedOnly
Assert-True ($ev.Data.planned) 'Planned event not persisted without logPlannedEventsInDryRun'

Import-Module "$repoRoot/modules/Processing/QC.CommentStatusProcessor.psm1" -Force
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force

# Invalid transition must not count as successful state apply when failOnStateApply=true
$settings = @{ failOnStateApply = $true }
$invalidOk = New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Message 'blocked' -Data @{}
Assert-True (-not (Test-QCCommentSyncStateApplySucceeded -StateRes $invalidOk -Settings $settings)) 'Soft-success invalid transition is treated as failure when failOnStateApply=true'

$settingsOff = @{ failOnStateApply = $false }
Assert-True (Test-QCCommentSyncStateApplySucceeded -StateRes $invalidOk -Settings $settingsOff) 'failOnStateApply=false allows legacy soft-success codes'

Write-Host 'OK test_qc_comment_sync_validation.ps1'
