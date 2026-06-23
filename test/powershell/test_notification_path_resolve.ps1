$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Notifications/QC.Notifications.psm1') -Force

$cfg = @{
    notifications = @{
        enabled = $true
        provider = 'Mock'
        dryRun = $true
        outputRoot = 'notifications'
        dedupe = @{
            enabled = $true
            storePath = 'notifications\dedupe\sent-keys.jsonl'
        }
    }
}
$s = Get-QCNotificationSettings -Config $cfg
$expectedRoot = Join-Path $repoRoot 'notifications'
$expectedStore = Join-Path $repoRoot 'notifications\dedupe\sent-keys.jsonl'
if ($s.outputRoot -ne $expectedRoot) {
    throw "outputRoot mismatch: $($s.outputRoot) != $expectedRoot"
}
if ($s.dedupe.storePath -ne $expectedStore) {
    throw "storePath mismatch: $($s.dedupe.storePath) != $expectedStore"
}
Write-Host 'test_notification_path_resolve: OK'
