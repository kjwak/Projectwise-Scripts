# Notification routing plan tests for comment sync.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/QC.CommentSync.Notifications.psm1" -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$config = @{
    qcCommentSync = @{
        targetStates = @{
            correctionsInProgress = 'Corrections In Progress'
            backcheckInProgress = 'Backcheck In Progress'
            completed = 'Corrections Complete'
            error = 'Error Needs Attention'
        }
    }
    notifications = @{ adminRecipients = @('admin@example.com') }
}

$meta = @{ fileName = 'A101-qc.pdf'; pwPath = 'Documents\X\A101-qc.pdf'; documentId = 'guid-1'; projectId = 'proj' }

$p1 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Corrections In Progress'; decisionCode = 'CORRECTIONS_REQUIRED'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p1.eventType 'CORRECTIONS_IN_PROGRESS' 'Corrections routes to designers event'
Assert-Eq $p1.toRoles[0] 'designers' 'To designers'

$p2 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Backcheck In Progress'; decisionCode = 'BACKCHECK_REQUIRED'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p2.eventType 'BACKCHECK_IN_PROGRESS' 'Backcheck event'
Assert-Eq $p2.toRoles[0] 'reviewers' 'To reviewers'

$p3 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Error Needs Attention'; decisionCode = 'PARSE_ERROR'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p3.eventType 'QC_ERROR' 'Error event'
Assert-Eq $p3.explicitTo[0] 'admin@example.com' 'Admin recipient on error'

Write-Host 'OK test_qc_comment_notification_routing.ps1'
