$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\QC.Notifications.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

Assert-Eq (Get-QCProcessTypePdfSuffix -ProcessType 'review' -Config @{}) 'rev' 'review suffix'
Assert-Eq (Get-QCProcessTypePdfSuffix -ProcessType 'production' -Config @{}) 'prod' 'production suffix'
Assert-Eq (Get-QCProcessTypePdfSuffix -ProcessType 'check' -Config @{}) 'chk' 'check suffix'

$norm = Normalize-QCProcessType -ProcessType 'Review' -ReviewType 'Production QC'
Assert-Eq $norm 'review' 'QC_Process_Type must win over QC_Review_Type'

$tmpStore = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_dedupe_test_' + [guid]::NewGuid().ToString('N') + '.jsonl')
$notifSettings = @{
    dedupe = @{
        enabled = $true
        storePath = $tmpStore
    }
}
Register-QCNotificationDedupe -DedupeKey 'test|pending' -Settings $notifSettings -ResultData @{
    eventType = 'READY_FOR_QC'
    documentName = 'sheet-prod.pdf'
    provider = 'Mock'
    status = 'pending'
}
if (Test-QCNotificationDedupe -DedupeKey 'test|pending' -Settings $notifSettings) {
    throw 'Pending dedupe claim must not block retries.'
}
Remove-QCNotificationDedupeKey -DedupeKey 'test|pending' -Settings $notifSettings | Out-Null
Register-QCNotificationDedupe -DedupeKey 'test|pending' -Settings $notifSettings -ResultData @{
    eventType = 'READY_FOR_QC'
    documentName = 'sheet-prod.pdf'
    provider = 'Mock'
    status = 'sent'
}
if (-not (Test-QCNotificationDedupe -DedupeKey 'test|pending' -Settings $notifSettings)) {
    throw 'Sent dedupe entry must block duplicates.'
}
Remove-Item -LiteralPath $tmpStore -Force -ErrorAction SilentlyContinue

Write-Host 'test_qc_notification_lane_pdf: OK' -ForegroundColor Green
