$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/QC.Notifications.psm1" -Force
Import-Module "$repoRoot/modules/QC.NotificationGraph.psm1" -Force

$testRoot = Join-Path $env:TEMP ("qc-notify-test-" + [guid]::NewGuid().ToString('N'))
$mockRoot = Join-Path $testRoot 'notifications'
$dedupePath = Join-Path $testRoot 'dedupe\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $dedupePath) -Force | Out-Null

function New-NotifyConfig([bool]$Enabled, [string]$Provider = 'Mock') {
    return @{
        notifications = @{
            enabled = $Enabled
            provider = $Provider
            dryRun = $true
            outputRoot = $mockRoot
            dedupe = @{
                enabled = $true
                storePath = $dedupePath
                keyFields = @('documentGuid', 'eventType', 'currentState')
            }
            graph = @{
                tenantId = ''
                clientId = ''
                senderMailbox = ''
                certificateThumbprint = ''
            }
            attributes = @{
                reviewerEmailField = 'EM_Reviewer_Email'
                designerEmailField = 'EM_Designer_Email'
                ccEmailField = 'CcEmails'
            }
            events = @{
                'QC Received' = @{
                    enabled = $true
                    eventType = 'QC_RECEIVED'
                    to = @('reviewers')
                    cc = @('designers')
                    subjectTemplate = 'QC Received - {documentName}'
                    actionRequired = 'Reviewer to begin QC review.'
                }
            }
        }
    }
}

function New-MockDocument([string]$Reviewer, [string]$Designer) {
    $bag = [System.Collections.Specialized.OrderedDictionary]::new()
    if ($Reviewer) { [void]$bag.Add('EM_Reviewer_Email', $Reviewer) }
    if ($Designer) { [void]$bag.Add('EM_Designer_Email', $Designer) }
    return [pscustomobject]@{
        Name = 'D-101-qc.pdf'
        DocumentGUID = 'guid-101'
        Attributes = @($bag)
    }
}

# Disabled: no mock files created
$disabledCfg = New-NotifyConfig -Enabled $false
$disabled = Invoke-QCNotificationForStateChange -Config $disabledCfg -PreviousState 'In Production' -CurrentState 'QC Received' -Document (New-MockDocument 'r@x.com' 'd@x.com')
Assert-True $disabled.IsSuccess 'Disabled notification should return success with skip'
Assert-True $disabled.Data.skipped 'Disabled notification should be skipped'
$mockFiles = @(Get-ChildItem -LiteralPath (Join-Path $mockRoot 'mock') -ErrorAction SilentlyContinue)
Assert-Eq $mockFiles.Count 0 'Disabled notifications must not write mock payloads'

# Enabled mock send
$enabledCfg = New-NotifyConfig -Enabled $true
$send = Invoke-QCNotificationForStateChange -Config $enabledCfg -PreviousState 'In Production' -CurrentState 'QC Received' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com') -Project 'CAFWY2200'
Assert-True $send.IsSuccess 'Mock notification should succeed'
Assert-True $send.Data.success 'Result success flag should be true'
Assert-Eq $send.Data.eventType 'QC_RECEIVED' 'State should map to QC_RECEIVED'
$mockFiles = @(Get-ChildItem -LiteralPath (Join-Path $mockRoot 'mock') -Filter '*.json')
Assert-True ($mockFiles.Count -ge 1) 'Mock provider should create a payload file'

# Missing recipients -> skipped/failure without crash
$noRecipientCfg = New-NotifyConfig -Enabled $true
$noRecipients = Invoke-QCNotificationForStateChange -Config $noRecipientCfg -PreviousState 'In Production' -CurrentState 'QC Received' `
    -Document (New-MockDocument '' '')
Assert-True ($null -ne $noRecipients) 'Missing recipients should return a result object'
Assert-True (-not $noRecipients.Data.success) 'Missing recipients should not report success'

# Dedupe key consistency
$event = New-QCNotificationEvent -EventType 'QC_RECEIVED' -DocumentName 'D-101-qc.pdf' -DocumentGuid 'guid-101' -CurrentState 'QC Received'
$settings = Get-QCNotificationSettings -Config $enabledCfg
$key1 = Get-QCNotificationDedupeKey -Event $event -Settings $settings
$key2 = Get-QCNotificationDedupeKey -Event $event -Settings $settings
Assert-Eq $key1 $key2 'Dedupe keys should be stable for the same event'

# Duplicate suppression
$dupCfg = New-NotifyConfig -Enabled $true
$dupDedupe = Join-Path $testRoot 'dedupe-dup\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $dupDedupe) -Force | Out-Null
$dupCfg.notifications.dedupe.storePath = $dupDedupe
$first = Invoke-QCNotificationForStateChange -Config $dupCfg -PreviousState 'In Production' -CurrentState 'QC Received' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com')
$second = Invoke-QCNotificationForStateChange -Config $dupCfg -PreviousState 'In Production' -CurrentState 'QC Received' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com')
Assert-True $second.Data.skipped 'Second identical notification should be deduped'
Assert-Eq $second.Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'Duplicate should use skip code'

# Graph not configured
$graphCfg = New-NotifyConfig -Enabled $true -Provider 'MicrosoftGraph'
$graph = Send-QCNotification -Event $event -Config $graphCfg -Subject 'Test' -To @('a@b.com')
Assert-True (-not $graph.IsSuccess) 'Graph without credentials should fail'
Assert-True ($graph.Message -match 'not configured') 'Graph should report not configured'

$graphValidation = Test-QCNotificationGraphConfigured -GraphSettings @{ tenantId = ''; clientId = ''; senderMailbox = '' }
Assert-True (-not $graphValidation.configured) 'Blank Graph settings should not be configured'

Write-Host 'All QC notification tests passed.'
