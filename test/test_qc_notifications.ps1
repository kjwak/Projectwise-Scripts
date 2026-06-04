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
                clientSecret = ''
                senderMailbox = ''
                certificateThumbprint = ''
            }
            attributes = @{
                reviewerEmailField = 'EM_Reviewer_Email'
                designerEmailField = 'EM_Designer_Email'
                ccEmailField = 'CcEmails'
            }
            email = @{
                subjectTemplate = '[{ReviewType}] {ProjectName} | {DocumentName} | {WorkflowState}'
            }
            events = @{
                'QC Received' = @{
                    enabled = $true
                    eventType = 'QC_RECEIVED'
                    to = @('reviewers')
                    cc = @('designers')
                    actionRequired = 'Reviewer to begin QC review.'
                }
            }
        }
    }
}

function New-MockDocument([string]$Reviewer, [string]$Designer, [string]$Checker = '', [string]$ReviewType = '') {
    $bag = [System.Collections.Specialized.OrderedDictionary]::new()
    if ($Reviewer) { [void]$bag.Add('EM_Reviewer_Email', $Reviewer) }
    if ($Designer) { [void]$bag.Add('EM_Designer_Email', $Designer) }
    if ($Checker) { [void]$bag.Add('EM_Checker_Email', $Checker) }
    if ($ReviewType) { [void]$bag.Add('QC_Review_Type', $ReviewType) }
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

# Independent Check routes reviewers role to checker email
$icCfg = New-NotifyConfig -Enabled $true
$icSettings = Get-QCNotificationSettings -Config $icCfg
$icDoc = New-MockDocument -Reviewer 'reviewer@company.com' -Designer 'designer@company.com' `
    -Checker 'checker@company.com' -ReviewType 'Independent Check'
$icResolved = Resolve-QCNotificationRecipients -Document $icDoc -Settings $icSettings -Config @{
    qcWorkflow = @{ reviewTypes = @{ independentCheck = 'Independent Check' } }
} -ToRoles @('reviewers')
Assert-Eq $icResolved.reviewers[0] 'checker@company.com' 'Independent Check should use checker for reviewers role'
Assert-Eq $icResolved.reviewers.Count 1 'Independent Check should not include EM_Reviewer in reviewers list'

# EM_* on document properties (audit rows without Attributes bag)
$emCfg = New-NotifyConfig -Enabled $true
$emCfg.notifications.events['Redlines Received'] = @{
    enabled = $true
    eventType = 'REDLINES_ISSUED'
    to = @('designers')
    cc = @('reviewers')
    subjectTemplate = 'Redlines - {documentName}'
    actionRequired = 'Designer to begin corrections.'
}
$emDoc = [pscustomobject]@{
    Name = '080J082001ab001-qc.pdf'
    DocumentGUID = 'guid-em-roles'
    FolderPath = 'Documents\AZDOT\TEST\CADD\Sheets'
    EM_Designer_Email = 'jflint@aztec.us'
    EM_Reviewer_Email = 'JFlint@aztec.us'
    EM_Checker_Email = 'jflint@aztec.us'
}
$emSend = Invoke-QCNotificationForStateChange -Config $emCfg -PreviousState 'Ready for QC' -CurrentState 'Redlines Received' `
    -Document $emDoc -DocumentPath 'Documents\AZDOT\TEST\CADD\Sheets\080J082001ab001-qc.pdf'
Assert-True $emSend.IsSuccess 'Notification with EM_* document properties should succeed'
Assert-True $emSend.Data.success 'EM_* roles should resolve to To recipients'

# QC_* workflow columns when EM_* absent
$qcOnlyCfg = New-NotifyConfig -Enabled $true
$qcOnlySettings = Get-QCNotificationSettings -Config $qcOnlyCfg
$qcResolved = Resolve-QCNotificationRecipients -Document ([pscustomobject]@{
    Name = 'sheet.pdf'
    QC_Designer_Email = 'designer@company.com'
    QC_Reviewer_Email = 'reviewer@company.com'
    QC_Checker_Email = 'checker@company.com'
}) -Settings $qcOnlySettings -Config @{
    qcWorkflow = @{
        attributeMap = @{
            designerEmail = 'QC_Designer_Email'
            reviewerEmail = 'QC_Reviewer_Email'
            checkerEmail = 'QC_Checker_Email'
        }
    }
} -ToRoles @('reviewers') -CcRoles @('designers')
Assert-Eq $qcResolved.reviewers[0] 'reviewer@company.com' 'QC_Reviewer_Email fallback should resolve'
Assert-Eq $qcResolved.designers[0] 'designer@company.com' 'QC_Designer_Email fallback should resolve'

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
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com') -StateTransitionKey 'audit:1001'
$second = Invoke-QCNotificationForStateChange -Config $dupCfg -PreviousState 'In Production' -CurrentState 'QC Received' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com') -StateTransitionKey 'audit:1001'
Assert-True $second.Data.skipped 'Second identical notification should be deduped'
Assert-Eq $second.Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'Duplicate should use skip code'

# New transition to same state should not dedupe (repeat QC cycle)
$cycleDedupe = Join-Path $testRoot 'dedupe-cycle\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $cycleDedupe) -Force | Out-Null
$cycleCfg = New-NotifyConfig -Enabled $true
$cycleCfg.notifications.dedupe.storePath = $cycleDedupe
$cycleCfg.notifications.events['Redlines Received'] = @{
    enabled = $true
    eventType = 'REDLINES_RECEIVED'
    to = @('designers')
    cc = @('reviewers')
    subjectTemplate = 'Redlines - {documentName}'
    actionRequired = 'Designer corrections.'
}
$docCycle = [pscustomobject]@{ Name = 'S1-qc.pdf'; DocumentGUID = 'guid-cycle'; EM_Designer_Email = 'd@x.com'; EM_Reviewer_Email = 'r@x.com' }
$r1 = Invoke-QCNotificationForStateChange -Config $cycleCfg -PreviousState 'Ready for QC' -CurrentState 'Redlines Received' `
    -Document $docCycle -StateTransitionKey 'audit:2001'
Assert-True $r1.IsSuccess 'First Redlines Received send should succeed'
$r2 = Invoke-QCNotificationForStateChange -Config $cycleCfg -PreviousState 'Corrections Received' -CurrentState 'Redlines Received' `
    -Document $docCycle -StateTransitionKey 'audit:2002'
Assert-True $r2.IsSuccess 'Second cycle Redlines Received should succeed with new transition key'
Assert-True (-not $r2.Data.skipped) 'Second cycle should not be deduped'

# Graph not configured
$graphCfg = New-NotifyConfig -Enabled $true -Provider 'MicrosoftGraph'
$graph = Send-QCNotification -Event $event -Config $graphCfg -Subject 'Test' -To @('a@b.com')
Assert-True (-not $graph.IsSuccess) 'Graph without credentials should fail'
Assert-True ($graph.Message -match 'not configured') 'Graph should report not configured'

$graphValidation = Test-QCNotificationGraphConfigured -GraphSettings @{ tenantId = ''; clientId = ''; senderMailbox = '' }
Assert-True (-not $graphValidation.configured) 'Blank Graph settings should not be configured'

# Graph configured with client secret + dryRun (no live API call)
$graphDryCfg = New-NotifyConfig -Enabled $true -Provider 'MicrosoftGraph'
$graphDryCfg.notifications.graph.tenantId = '00000000-0000-0000-0000-000000000001'
$graphDryCfg.notifications.graph.clientId = '00000000-0000-0000-0000-000000000002'
$graphDryCfg.notifications.graph.clientSecret = 'test-secret'
$graphDryCfg.notifications.graph.senderMailbox = 'qc-notify@example.com'
$graphDry = Send-QCNotification -Event $event -Config $graphDryCfg -Subject 'Test' -To @('a@b.com') -Body 'Body'
Assert-True $graphDry.IsSuccess 'Graph dry run with client secret should succeed'
Assert-True $graphDry.Data.dryRun 'Graph dry run should set dryRun flag'
Assert-Eq $graphDry.Code 'QC_NOTIFICATION_GRAPH_DRY_RUN' 'Graph dry run code'

Write-Host 'All QC notification tests passed.'
