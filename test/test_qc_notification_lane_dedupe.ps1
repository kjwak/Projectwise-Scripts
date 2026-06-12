# Notification dedupe keys include qc_process_type so lanes do not collide.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Notifications.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Ne($a, $b, $msg) { if ($a -eq $b) { throw "ASSERT FAILED: $msg" } }

$settings = Get-QCNotificationSettings -Config @{}
Assert-True ($settings.dedupe.keyFields -contains 'qcProcessType') 'default dedupe includes qcProcessType'

$baseEvent = @{
    sheetStem = 'CA001'
    previousState = 'QC Initiated'
    currentState = 'Ready for QC'
    transitionSource = 'user_audit'
    logicalTransitionAnchor = 'CA001|Ready for QC'
    recipientKey = 'a@example.com'
    folderPath = 'Documents\X\CADD\Sheets'
}

$prodEvent = $baseEvent.Clone()
$prodEvent['qcProcessType'] = 'production'
$prodEvent['documentName'] = 'CA001-prod.pdf'
$prodEvent['documentGuid'] = '11111111-1111-1111-1111-111111111111'

$chkEvent = $baseEvent.Clone()
$chkEvent['qcProcessType'] = 'check'
$chkEvent['documentName'] = 'CA001-chk.pdf'
$chkEvent['documentGuid'] = '22222222-2222-2222-2222-222222222222'

$revEvent = $baseEvent.Clone()
$revEvent['qcProcessType'] = 'review'
$revEvent['documentName'] = 'CA001-rev.pdf'
$revEvent['documentGuid'] = '33333333-3333-3333-3333-333333333333'

$prodKey = Get-QCNotificationDedupeKey -Event $prodEvent -Settings $settings
$chkKey = Get-QCNotificationDedupeKey -Event $chkEvent -Settings $settings
$revKey = Get-QCNotificationDedupeKey -Event $revEvent -Settings $settings

Assert-Ne $prodKey $chkKey 'production and check dedupe keys differ'
Assert-Ne $prodKey $revKey 'production and review dedupe keys differ'
Assert-Ne $chkKey $revKey 'check and review dedupe keys differ'

Write-Host 'test_qc_notification_lane_dedupe: OK' -ForegroundColor Green
