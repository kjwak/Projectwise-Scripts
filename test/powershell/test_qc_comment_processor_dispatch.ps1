# Verify QC_COMMENT_STATUS_SYNC maps to orchestrator handler.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$job = @{
    id = 'j1'
    type = 'QC_COMMENT_STATUS_SYNC'
    sourcePath = 'Documents\X\A101-qc.pdf'
    sourceName = 'A101-qc.pdf'
    sourceFolder = 'Documents\X'
    triggerRule = @{ id = 'qc-comment-status-pw'; jobType = 'QC_COMMENT_STATUS_SYNC' }
    dedupeKey = 'dq_test'
    metadata = @{}
}

$config = @{
    dryRun = $true
    qcCommentSync = @{ enabled = $true; stagingRoot = $env:TEMP }
    processors = @{
        processorMap = @{ QC_COMMENT_STATUS_SYNC = 'Invoke-QCCommentStatusSyncProcessor' }
    }
}

$ready = Test-QCJobReady -Job $job -Config $config
Assert-True $ready.IsSuccess 'Job should be ready'

$cmd = Get-Command -Name 'Invoke-QCCommentStatusSyncProcessor' -ErrorAction SilentlyContinue
Assert-True ($null -ne $cmd) 'Comment sync processor should be loaded via QC.Processors'

Write-Host 'OK test_qc_comment_processor_dispatch.ps1'
