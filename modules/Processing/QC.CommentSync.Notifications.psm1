# QC.CommentSync.Notifications.psm1
# Responsibility: Route comment-sync notifications by decision target state.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.PdfExport.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.CommentStatusDecision.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.Notifications.psm1') -Force -ErrorAction SilentlyContinue

function Get-QCCommentSyncNotificationPlan {
    param(
        [hashtable]$Config,
        [hashtable]$Decision,
        [hashtable]$JobMetadata
    )

    $settings = Get-QCCommentDecisionSettings -Config $Config
    $targets = $settings.targetStates
    $targetState = [string]$Decision.targetState

    $eventType = 'QC_COMMENT_SYNC'
    $toRoles = @()
    $ccRoles = @('reviewers')
    $explicitTo = @()

    if ($targetState -eq [string]$targets.redlinesReceived -or $targetState -eq [string]$targets.redlinesIssued) {
        $eventType = 'REDLINES_RECEIVED'
        $toRoles = @('designers')
        $ccRoles = @('reviewers')
    } elseif ($targetState -eq [string]$targets.qcFinalizing) {
        $eventType = 'QC_FINALIZING'
        $toRoles = @()
        $ccRoles = @()
    } elseif ($targetState -eq [string]$targets.error) {
        $eventType = 'QC_ERROR'
        $toRoles = @()
        $ccRoles = @()
        if ($Config.notifications -and $Config.notifications.adminRecipients) {
            $explicitTo = @($Config.notifications.adminRecipients | ForEach-Object { [string]$_ })
        }
    } else {
        $eventType = 'QC_COMMENT_SYNC_COMPLETE'
        $toRoles = @('reviewers')
        $ccRoles = @('designers')
    }

    return @{
        eventType = $eventType
        toRoles = $toRoles
        ccRoles = $ccRoles
        explicitTo = $explicitTo
        targetState = $targetState
        decisionCode = [string]$Decision.decisionCode
        summary = [string]$Decision.summary
        documentName = [string]$JobMetadata.fileName
        documentPath = [string]$JobMetadata.pwPath
        documentGuid = [string]$JobMetadata.documentId
    }
}

function Send-QCCommentSyncNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$Decision,
        [Parameter(Mandatory)]
        [hashtable]$JobMetadata,
        [object]$Document = $null,
        [hashtable]$Job = $null,
        [string]$PreviousState = ''
    )

    $plan = Get-QCCommentSyncNotificationPlan -Config $Config -Decision $Decision -JobMetadata $JobMetadata
    $dryRun = $false
    if ($Config.ContainsKey('dryRun')) { try { $dryRun = [bool]$Config.dryRun } catch { } }
    if ($Config.notifications -and $null -ne $Config.notifications.dryRun) {
        try { if ([bool]$Config.notifications.dryRun) { $dryRun = $true } } catch { }
    }

    if (-not (Get-Command -Name 'Send-QCNotification' -ErrorAction SilentlyContinue)) {
        return New-QCSuccessResult -Code 'QC_COMMENT_NOTIFY_PLANNED' -Message 'Notifications module unavailable; planned only.' -Data @{ planned = $true; plan = $plan }
    }

    $notifSettings = $null
    if (Get-Command -Name 'Get-QCNotificationSettings' -ErrorAction SilentlyContinue) {
        $notifSettings = Get-QCNotificationSettings -Config $Config
    }

    $resolved = @{ to = @(); cc = @() }
    if (Get-Command -Name 'Resolve-QCNotificationRecipients' -ErrorAction SilentlyContinue) {
        $resolved = Resolve-QCNotificationRecipients -Document $Document -Settings $notifSettings `
            -ToRoles @($plan.toRoles) -CcRoles @($plan.ccRoles) -ExplicitCc @()
    }
    $to = @($resolved.to)
    if (@($plan.explicitTo).Count -gt 0) { $to = @($plan.explicitTo) }

    $event = @{
        eventType = $plan.eventType
        project = if ($JobMetadata.projectId) { [string]$JobMetadata.projectId } else { '' }
        documentName = $plan.documentName
        documentPath = $plan.documentPath
        documentGuid = $plan.documentGuid
        previousState = $PreviousState
        currentState = $plan.targetState
        actionRequired = $plan.summary
        sourceJobId = if ($Job -and $Job.id) { [string]$Job.id } else { '' }
    }

    if ($dryRun) {
        return New-QCSuccessResult -Code 'QC_COMMENT_NOTIFY_PLANNED' -Message 'Dry-run: notification not sent.' -Data @{
            planned = $true
            event = $event
            to = $to
            cc = @($resolved.cc)
        }
    }

    try {
        $send = Send-QCNotification -Config $Config -Event $event -To $to -Cc @($resolved.cc)
        return $send
    } catch {
        return New-QCFailureResult -Code 'QC_COMMENT_NOTIFY_FAILED' -Message $_.Exception.Message -Data @{ plan = $plan }
    }
}

Export-ModuleMember -Function Get-QCCommentSyncNotificationPlan, Send-QCCommentSyncNotification
