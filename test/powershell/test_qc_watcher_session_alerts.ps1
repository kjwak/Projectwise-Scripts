$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Core/Core.Runtime.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.NotificationGraph.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.NotificationMock.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.WatcherAlerts.psm1" -Force

$testRoot = Join-Path $env:TEMP ("qc-watcher-alert-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

$config = @{
    dryRun = $true
    notifications = @{
        enabled = $true
        provider = 'Mock'
        dryRun = $true
        outputRoot = (Join-Path $testRoot 'notifications')
        graph = @{
            tenantId = ''
            clientId = ''
            clientSecret = ''
            senderMailbox = ''
        }
    }
    watcher = @{
        sessionAlerts = @{
            enabled = $true
            recipients = @('jflint@aztec.us')
            importance = 'high'
            dedupeMinutes = 60
            probeIntervalTicks = 30
            staleAuditLagSeconds = 180
        }
    }
}

$settings = Get-QCWatcherSessionAlertSettings -Config $config
Assert-Eq $settings.recipients[0] 'jflint@aztec.us' 'default recipient from config'
Assert-Eq $settings.importance 'high' 'importance should be high'

$due1 = Test-QCWatcherSessionAlertDue -Config $config
Assert-True $due1.due 'first alert should be due'

$send1 = Send-QCWatcherSessionLostAlert -Config $config -Details @{
    detectedUtc = (Get-Date).ToUniversalTime().ToString('o')
    reason = 'health_probe_failed'
    datasourceName = 'test-ds'
    tick = '42'
    errorMessage = 'probe failed'
}
Assert-True $send1.IsSuccess "alert send should succeed: $($send1.Message)"

$due2 = Test-QCWatcherSessionAlertDue -Config $config
Assert-True (-not $due2.due) 'second alert should be deduped within window'

Import-Module "$repoRoot/modules/Notifications/QC.NotificationGraph.psm1" -Force
$graphMsg = New-QCGraphEmailMessage -ToRecipients @('jflint@aztec.us') -Subject 'Test' -HtmlBody '<p>hi</p>' `
    -LogoPath (Join-Path $repoRoot 'email/typsalogo.png.webp') -Importance 'high'
Assert-Eq $graphMsg.message.importance 'high' 'Graph message should carry high importance'

$stallDisabled = Test-QCWatcherAuditActivityStalled -Config $config `
    -MaxPwActTimeUtc '2026-06-10 07:17:22' `
    -PollUntilUtc '2026-06-10 07:24:00' `
    -LastMaxPwActChangeUtc ([datetime]'2026-06-10T07:17:22Z')
Assert-True (-not $stallDisabled.stalled) 'audit stall detection should be disabled by default'

$config.watcher.sessionAlerts.enableAuditStallDetection = $true
$stallEnabled = Test-QCWatcherAuditActivityStalled -Config $config `
    -MaxPwActTimeUtc '2026-06-10 07:17:22' `
    -PollUntilUtc '2026-06-10 07:24:00' `
    -LastMaxPwActChangeUtc ([datetime]'2026-06-10T07:17:22Z')
Assert-True $stallEnabled.stalled 'audit stall heuristic should detect lag when explicitly enabled'

Write-Host 'OK: QC watcher session alert tests passed.' -ForegroundColor Green
