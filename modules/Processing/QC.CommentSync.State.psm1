# QC.CommentSync.State.psm1
# Responsibility: Apply workflow state to *-qc.pdf via existing QC.Workflow module.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.Workflow.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.PdfExport.psm1') -Force

function Test-QCCommentSyncDryRun {
    param([hashtable]$Config)
    if ($Config -and $Config.ContainsKey('dryRun') -and [bool]$Config.dryRun) { return $true }
    $wf = $null
    if ($Config -and $Config.qcWorkflow) { $wf = $Config.qcWorkflow }
    if ($wf -and $wf.dryRunWriteback -eq $true) { return $true }
    return $false
}

function Set-QCPdfCommentSyncWorkflowState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$TargetStateName,
        [object]$Document,
        [hashtable]$Job = $null,
        [string]$PreviousStateName = '',
        [switch]$DryRun
    )

    $isDryRun = [bool]$DryRun
    if (-not $isDryRun) { $isDryRun = Test-QCCommentSyncDryRun -Config $Config }

    if ($isDryRun) {
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_PLANNED' -Message 'Dry-run: would set document workflow state.' -Data @{
            planned = $true
            stateName = $TargetStateName
            previousStateName = $PreviousStateName
            changed = $false
        }
    }

    if (-not (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_COMMENT_STATE_MODULE_MISSING' -Message 'QC.Workflow not loaded.' -Data @{
            planned = $true
            targetStateName = $TargetStateName
            previousStateName = $PreviousStateName
        }
    }

    if ([string]::IsNullOrWhiteSpace($TargetStateName)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_EMPTY_STATE_GUARDED' `
                -Message 'Skipped comment-sync workflow state write because target StateName was empty.' -Data @{
                callSite = 'Set-QCPdfCommentSyncWorkflowState.TargetStateName'
                auditEventId = $null
                documentName = if ($Job -and $Job.ContainsKey('sourceName')) { [string]$Job.sourceName } else { '' }
                folderPath = if ($Job -and $Job.ContainsKey('sourceFolder')) { [string]$Job.sourceFolder } else { '' }
                sourceVariableName = 'TargetStateName'
                sourceValue = $TargetStateName
                livePwState = $PreviousStateName
                changedByUsername = ''
            } | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_COMMENT_STATE_EMPTY_TARGET_SKIPPED' -Message 'Comment-sync target workflow state was empty; no state change was made.' -Data @{
            planned = $false; changed = $false; targetStateName = $TargetStateName; previousStateName = $PreviousStateName
        }
    }

    $settings = Get-QCWorkflowSettings -Config $Config
    $context = @{
        document = $Document
        job = $Job
        previousState = $PreviousStateName
    }

    $result = Set-PWQCWorkflowState -Settings $settings -Context $context -StateName $TargetStateName -DryRun:$isDryRun
  return $result
}

Export-ModuleMember -Function Test-QCCommentSyncDryRun, Set-QCPdfCommentSyncWorkflowState
