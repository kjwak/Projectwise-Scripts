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
Import-Module "$repoRoot/modules/QC.NotificationTemplates.psm1" -Force
Import-Module "$repoRoot/modules/QC.NotificationGraph.psm1" -Force
Import-Module "$repoRoot/modules/QC.AuditTriggers.psm1" -Force

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
                keyFields = @('sheetStem', 'eventType', 'previousState', 'currentState')
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

# Missing email attributes for configured state -> field list
$missingCfg = New-NotifyConfig -Enabled $true
$missingCfg.notifications.rollbackWhenEmailAttributesMissing = $true
$missingCfg.notifications.events['Ready for QC'] = @{
    enabled = $true
    eventType = 'READY_FOR_QC'
    to = @('reviewers')
    cc = @('designers')
}
$missingFields = @(Get-QCStateChangeMissingEmailFields -Config $missingCfg -TargetStateName 'Ready for QC' `
    -Document (New-MockDocument '' ''))
Assert-True ($missingFields.Count -ge 2) 'Should report missing designer and reviewer fields'
Assert-True ($missingFields -contains 'EM_Designer_Email') 'Should include EM_Designer_Email'
Assert-True ($missingFields -contains 'EM_Reviewer_Email') 'Should include EM_Reviewer_Email'
$okFields = @(Get-QCStateChangeMissingEmailFields -Config $missingCfg -TargetStateName 'Ready for QC' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com'))
Assert-Eq $okFields.Count 0 'Populated emails should yield no missing fields'

# QC Initiated also requires emails for post-prepend Ready for QC notification
$initiatedMissing = @(Get-QCWorkflowTransitionMissingEmailFields -Config $missingCfg -TargetStateName 'QC Initiated' `
    -Document (New-MockDocument '' ''))
Assert-True ($initiatedMissing.Count -ge 2) 'QC Initiated should inherit Ready for QC email requirements'
$prependBlock = Test-QCPrependBlockedByMissingEmailAttributes -Config $missingCfg -SheetPdfName 'sheet001.pdf'
Assert-True $prependBlock.blocked 'Prepend should be blocked when notification emails are missing'
$prependOk = Test-QCPrependBlockedByMissingEmailAttributes -Config $missingCfg -SheetPdfName 'sheet001.pdf' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com')
Assert-True (-not $prependOk.blocked) 'Prepend should proceed when notification emails are present'

# Rollback previous state falls back to production when sheet_index would match the blocked target
$rollbackCfg = New-NotifyConfig -Enabled $true
$rollbackCfg.qcWorkflow = @{ states = @{ production = 'In Production'; qcInitiated = 'QC Initiated'; readyForQc = 'Ready for QC' } }
$rollbackPrevious = Resolve-QCWorkflowRollbackPreviousState -Config $rollbackCfg -TargetStateName 'QC Initiated' `
    -Members @(@{ documentName = 'sheet001.pdf'; documentGuid = 'guid-sheet' }) -SheetPdfGuid 'guid-sheet'
Assert-Eq $rollbackPrevious 'In Production' 'Rollback should fall back to production when index has no prior state'

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

# Placeholder document names should not fork dedupe when folder + sheet stem align
$folderDedupeCfg = New-NotifyConfig -Enabled $true
$folderDedupeCfg.notifications.dedupe.keyFields = @('transitionId', 'stateTransitionKey', 'folderPath', 'sheetStem', 'eventType', 'previousState', 'currentState')
$folderDedupeSettings = Get-QCNotificationSettings -Config $folderDedupeCfg
$sharedFolder = 'documents\caltrans\proj\cadd\sheets\seg_1'
$keyUnknown = Get-QCNotificationDedupeKey -Event @{
    eventType = 'READY_FOR_QC'
    documentName = 'unknown-document-qc.pdf'
    sheetStem = '080J082001ab001'
    folderPath = $sharedFolder
    previousState = 'QC Initiated'
    currentState = 'Ready for QC'
    transitionId = 77
    stateTransitionKey = 'audit:5001'
} -Settings $folderDedupeSettings
$keyNamed = Get-QCNotificationDedupeKey -Event @{
    eventType = 'READY_FOR_QC'
    documentName = '080J082001ab001-qc.pdf'
    sheetStem = '080J082001ab001'
    folderPath = $sharedFolder
    previousState = 'QC Initiated'
    currentState = 'Ready for QC'
    transitionId = 77
    stateTransitionKey = 'audit:5001'
} -Settings $folderDedupeSettings
Assert-Eq $keyUnknown $keyNamed 'Same folder + sheet stem should dedupe despite different document names'
Assert-True ($keyUnknown -match 'folderPath=') 'Dedupe key should include folderPath'

# Subject should use sheet PDF name (stem.pdf), not unknown-document placeholder from QC PDF resolution
$readyCfg = New-NotifyConfig -Enabled $true
$readyDedupe = Join-Path $testRoot 'dedupe-ready-subject\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $readyDedupe) -Force | Out-Null
$readyCfg.notifications.dedupe.storePath = $readyDedupe
$readyCfg.notifications.events['Ready for QC'] = @{
    enabled = $true
    eventType = 'READY_FOR_QC'
    to = @('reviewers')
    cc = @('designers')
    actionRequired = 'Begin QC review.'
}
$readyJob = @{
    id = 'qc_notification_subject_test'
    sourceName = '080J082001ca001.pdf'
    sourceFolder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
    metadata = @{
        previousState = 'QC Initiated'
        currentState = 'Ready for QC'
        documentGuid = 'guid-ready-subject'
        stateTransitionKey = 'audit:9999'
        transitionId = 88
    }
}
$readyDoc = [pscustomobject]@{
    Name = '080J082001ca001.pdf'
    DocumentGUID = 'guid-ready-subject'
    EM_Reviewer_Email = 'reviewer@company.com'
    EM_Designer_Email = 'designer@company.com'
}
$readySend = Invoke-QCNotificationForStateChange -Config $readyCfg -PreviousState 'QC Initiated' -CurrentState 'Ready for QC' `
    -Document $readyDoc -Job $readyJob -DocumentGuid 'guid-ready-subject' -StateTransitionKey 'audit:9999' `
    -TransitionId 88 -Project 'cafwy2200-i-15_elpse'
Assert-True $readySend.IsSuccess 'Ready for QC job notification should succeed'
$readyPayload = Get-Content -LiteralPath $readySend.Data.filePath -Raw | ConvertFrom-Json
Assert-True ($readyPayload.subject -match '080J082001ca001\.pdf') 'Subject should use sheet PDF name'
Assert-True ($readyPayload.subject -notmatch 'unknown-document') 'Subject must not contain unknown-document placeholder'

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

# Different audit ids for the same sheet transition should still dedupe (echo audits / sibling sync)
$auditEchoDedupe = Join-Path $testRoot 'dedupe-audit-echo\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $auditEchoDedupe) -Force | Out-Null
$auditEchoCfg = New-NotifyConfig -Enabled $true
$auditEchoCfg.notifications.dedupe.storePath = $auditEchoDedupe
$auditEchoCfg.notifications.events['QC Complete'] = @{
    enabled = $true
    eventType = 'QC_COMPLETE'
    to = @('designers')
    cc = @('reviewers')
    actionRequired = 'QC review cycle is complete.'
}
$auditEchoDoc = [pscustomobject]@{
    Name = '081800063ea515-qc.pdf'
    DocumentGUID = 'guid-qc-complete'
    EM_Designer_Email = 'designer@company.com'
    EM_Reviewer_Email = 'reviewer@company.com'
}
$ae1 = Invoke-QCNotificationForStateChange -Config $auditEchoCfg -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -Document $auditEchoDoc -StateTransitionKey 'audit:9001'
Assert-True $ae1.IsSuccess 'First QC Complete send should succeed'
$ae2 = Invoke-QCNotificationForStateChange -Config $auditEchoCfg -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -Document $auditEchoDoc -StateTransitionKey 'audit:9002'
Assert-True $ae2.Data.skipped 'Second audit echo for same sheet transition should dedupe'
Assert-Eq $ae2.Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'Audit echo duplicate should use skip code'

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

# Transition id drives dedupe key and suppresses echo retries for the same transition
$tidDedupe = Join-Path $testRoot 'dedupe-transition-id\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $tidDedupe) -Force | Out-Null
$tidCfg = New-NotifyConfig -Enabled $true
$tidCfg.notifications.dedupe.storePath = $tidDedupe
$tidCfg.notifications.events['QC Complete'] = @{
    enabled = $true
    eventType = 'QC_COMPLETE'
    to = @('designers')
    cc = @('reviewers')
    actionRequired = 'QC review cycle is complete.'
}
$tidSettings = Get-QCNotificationSettings -Config $tidCfg
$tidEvent = @{ eventType = 'QC_COMPLETE'; documentName = '081800063ea515-qc.pdf'; sheetStem = '081800063ea515'; previousState = 'QC Finalizing'; currentState = 'QC Complete'; transitionId = 42 }
$tidKey = Get-QCNotificationDedupeKey -Event $tidEvent -Settings $tidSettings
Assert-True ($tidKey -match 'sheetStem=081800063ea515') 'Dedupe key should include logical sheet stem'
$tidDoc = [pscustomobject]@{
    Name = '081800063ea515-qc.pdf'
    DocumentGUID = 'guid-tid'
    EM_Designer_Email = 'designer@company.com'
    EM_Reviewer_Email = 'reviewer@company.com'
}
$t1 = Invoke-QCNotificationForStateChange -Config $tidCfg -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -Document $tidDoc -StateTransitionKey 'audit:9101' -TransitionId 42
Assert-True $t1.IsSuccess 'First send with transition id should succeed'
$t2 = Invoke-QCNotificationForStateChange -Config $tidCfg -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -Document $tidDoc -StateTransitionKey 'audit:9102' -TransitionId 42
Assert-True $t2.Data.skipped 'Same transition id with different audit echo should dedupe'
$t3 = Invoke-QCNotificationForStateChange -Config $tidCfg -PreviousState 'Corrections Received' -CurrentState 'QC Complete' `
    -Document $tidDoc -StateTransitionKey 'audit:9103' -TransitionId 43
Assert-True $t3.IsSuccess 'New transition id (QC cycle) should send again'
Assert-True (-not $t3.Data.skipped) 'New transition id should not be deduped'

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

# Global email subjectTemplate wins over legacy per-event subjectTemplate
$legacySubjectCfg = @{
    notifications = @{
        enabled = $true
        email = @{
            subjectTemplate = '[{ReviewType}] {ProjectName} | {DocumentName} | {WorkflowState}'
        }
        events = @{
            'Corrections Received' = @{
                enabled = $true
                eventType = 'CORRECTIONS_RECEIVED'
                subjectTemplate = 'Corrections Received - {documentName}'
                to = @('reviewers')
            }
        }
    }
}
$legacySettings = Get-QCNotificationSettings -Config $legacySubjectCfg
$legacyEventCfg = $legacySettings.events['Corrections Received']
Assert-True (-not ($legacyEventCfg -is [hashtable] -and $legacyEventCfg.ContainsKey('subjectTemplate'))) `
    'legacy per-event subjectTemplate should be stripped when email.subjectTemplate is set'
$legacyTemplate = [string]$legacySettings.email.subjectTemplate
$legacyTokens = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
$legacyTokens['documentName'] = '0818000063ea501-qc.pdf'
$legacyTokens['DocumentName'] = '0818000063ea501-qc.pdf'
$legacyTokens['project'] = 'cafwy2200'
$legacyTokens['ProjectName'] = 'cafwy2200'
$legacyTokens['currentState'] = 'Corrections Received'
$legacyTokens['WorkflowState'] = 'Corrections Received'
$legacyTokens['reviewType'] = 'Production QC'
$legacyTokens['ReviewType'] = 'Production QC'
$legacySubject = Expand-QCNotificationTemplate -Template $legacyTemplate -Tokens $legacyTokens
Assert-Eq $legacySubject '[Production QC] cafwy2200 | 0818000063ea501-qc.pdf | Corrections Received' `
    'subject should use global template format'

# SubmittedBy resolves state-change actor, not designer email
$actorCfg = @{
    auditPoller = @{
        workflowTriggers = @{
            automationPwUsernames = @('SVC_TYPSA_ARCHIVIST')
            automationPwUserNumbers = @(42)
        }
    }
}
Assert-Eq (Resolve-QCNotificationSubmittedBy -Config $actorCfg -ChangedByUsername 'jflint@typsa.com') `
    'jflint@typsa.com' 'explicit username should win'
Assert-Eq (Resolve-QCNotificationSubmittedBy -Config $actorCfg -ChangedByUser 42 -ChangedByUsername 'SVC_TYPSA_ARCHIVIST') `
    'SVC_TYPSA_ARCHIVIST' 'automation actor should use service username'
Assert-Eq (Resolve-QCNotificationSubmittedBy -Config $actorCfg -ChangedByUser 99) 'PW User 99' 'unknown user falls back to PW User n'

# SubmittedBy display is resolved from the same PW userno as transition_events.changed_by_user
$actorJob = @{
    metadata = @{
        changedByUser = 99
        changedByUsername = 'audit.user'
        lastActionBy = 'designer@typsa.com'
    }
}
$resolvedActor = Resolve-QCNotificationStateChangeActor -StateTransitionKey '' -ChangedByUser $null `
    -ChangedByUsername '' -Job $actorJob
Assert-Eq $resolvedActor.changedByUser 99 'job metadata changedByUser is used'
Assert-Eq $resolvedActor.changedByUsername 'audit.user' 'changedByUsername only from metadata, not designer fallbacks'
$display = Resolve-QCNotificationSubmittedBy -Config $actorCfg -ChangedByUser $resolvedActor.changedByUser `
    -ChangedByUsername $resolvedActor.changedByUsername
Assert-Eq $display 'PW User 99' 'SubmittedBy follows resolved changedByUser'

Write-Host 'All QC notification tests passed.'
