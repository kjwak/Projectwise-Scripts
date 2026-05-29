# Notification routing plan tests for comment sync.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/QC.CommentSync.Notifications.psm1" -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$config = @{
    qcCommentSync = @{
        targetStates = @{
            redlinesIssued = 'Redlines Issued'
            correctionsInProgress = 'Corrections In Progress'
            verificationInProgress = 'Verification In Progress'
            completed = 'QC Complete'
            error = 'Error Needs Attention'
        }
    }
    notifications = @{ adminRecipients = @('admin@example.com') }
}

$meta = @{ fileName = 'A101-qc.pdf'; pwPath = 'Documents\X\A101-qc.pdf'; documentId = 'guid-1'; projectId = 'proj' }

$p0 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Redlines Issued'; decisionCode = 'REDLINES_ISSUED'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p0.eventType 'REDLINES_ISSUED' 'Redlines issued event'
Assert-Eq $p0.toRoles[0] 'designers' 'Redlines issued to designers'

$p1 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Corrections In Progress'; decisionCode = 'CORRECTIONS_REQUIRED'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p1.eventType 'CORRECTIONS_IN_PROGRESS' 'Corrections routes to designers event'
Assert-Eq $p1.toRoles[0] 'designers' 'To designers'

$p2 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Verification In Progress'; decisionCode = 'VERIFICATION_REQUIRED'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p2.eventType 'VERIFICATION_IN_PROGRESS' 'Verification event'
Assert-Eq $p2.toRoles[0] 'reviewers' 'To reviewers'

$p3 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Error Needs Attention'; decisionCode = 'PARSE_ERROR'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p3.eventType 'QC_ERROR' 'Error event'
Assert-Eq $p3.explicitTo[0] 'admin@example.com' 'Admin recipient on error'

Write-Host 'OK test_qc_comment_notification_routing.ps1'
