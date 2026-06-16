# Notification routing plan tests for comment sync.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/QC.CommentSync.Notifications.psm1" -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$config = @{
    qcCommentSync = @{
        targetStates = @{
            redlinesReceived = 'Redlines Received'
            qcFinalizing = 'Initiate Verification'
            completed = 'Verified'
            error = 'Error Needs Attention'
        }
    }
    notifications = @{ adminRecipients = @('admin@example.com') }
}

$meta = @{ fileName = 'A101-qc.pdf'; pwPath = 'Documents\X\A101-qc.pdf'; documentId = 'guid-1'; projectId = 'proj' }

$p0 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Redlines Received'; decisionCode = 'REDLINES_RECEIVED'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p0.eventType 'REDLINES_RECEIVED' 'Redlines received event'
Assert-Eq $p0.toRoles[0] 'designers' 'Redlines received to designers'

$p1 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Initiate Verification'; decisionCode = 'QC_FINALIZING'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p1.eventType 'QC_FINALIZING' 'Initiate verification routes to automation intake (no email)'
Assert-Eq $p1.toRoles.Count 0 'No recipients on initiate verification'

$p2 = Get-QCCommentSyncNotificationPlan -Config $config -Decision @{ targetState = 'Error Needs Attention'; decisionCode = 'PARSE_ERROR'; summary = 'x' } -JobMetadata $meta
Assert-Eq $p2.eventType 'QC_ERROR' 'Error event'
Assert-Eq $p2.explicitTo[0] 'admin@example.com' 'Admin recipient on error'

Write-Host 'OK test_qc_comment_notification_routing.ps1'
