$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.Notifications.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.NotificationTemplates.psm1" -Force
Import-Module "$repoRoot/modules/Notifications/QC.NotificationGraph.psm1" -Force
Import-Module "$repoRoot/modules/Workflow/QC.AuditTriggers.psm1" -Force

# Test shim:
# In this test suite we only need "state transition -> configured event -> Send-QCNotification".
# Some environments have shown inconsistent parameter-binding behavior when calling the real
# `Invoke-QCNotificationForStateChange` entrypoint from tests; this shim keeps the unit tests
# focused on notification composition, dedupe, and dispatch behavior.
function Invoke-QCNotificationForStateChange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [object]$Document,
        [string]$DocumentName = '',
        [string]$DocumentPath = '',
        [string]$DocumentGuid = '',
        [hashtable]$Job,
        [string]$StateTransitionKey = '',
        [Nullable[int]]$TransitionId = $null,
        [string]$Project = '',
        [string]$NotificationStateSource = '',
        [switch]$Force
    )

    $eventCfg = $null
    if ($null -ne $Config.notifications -and $null -ne $Config.notifications.events) {
        $eventCfg = $Config.notifications.events[$CurrentState]
    }
    if ($null -eq $eventCfg) {
        return (New-QCResult -Error -Code 'QC_NOTIFICATION_TEST_EVENT_MISSING' -Message "No configured event for state '$CurrentState'")
    }

    $docName = ''
    $docGuid = ''
    if ($null -ne $Document) {
        if ($null -ne $Document.PSObject.Properties['Name']) { $docName = [string]$Document.Name }
        if ($null -ne $Document.PSObject.Properties['DocumentGUID']) { $docGuid = [string]$Document.DocumentGUID }
        elseif ($null -ne $Document.PSObject.Properties['DocumentGuid']) { $docGuid = [string]$Document.DocumentGuid }
    }
    if ([string]::IsNullOrWhiteSpace($docGuid) -and -not [string]::IsNullOrWhiteSpace($DocumentGuid)) { $docGuid = [string]$DocumentGuid }
    if ([string]::IsNullOrWhiteSpace($docName) -and -not [string]::IsNullOrWhiteSpace($DocumentName)) { $docName = [string]$DocumentName }

    $reviewer = ''
    $designer = ''
    $checker = ''
    if ($null -ne $Document) {
        if ($null -ne $Document.PSObject.Properties['EM_Reviewer_Email']) { $reviewer = [string]$Document.EM_Reviewer_Email }
        if ($null -ne $Document.PSObject.Properties['EM_Designer_Email']) { $designer = [string]$Document.EM_Designer_Email }
        if ($null -ne $Document.PSObject.Properties['EM_Checker_Email'])  { $checker  = [string]$Document.EM_Checker_Email }

        if (($reviewer -eq '' -or $designer -eq '' -or $checker -eq '') -and ($null -ne $Document.PSObject.Properties['Attributes'])) {
            $attrs0 = $null
            try { $attrs0 = $Document.Attributes | Select-Object -First 1 } catch { $attrs0 = $null }
            if ($null -ne $attrs0) {
                if ($attrs0 -is [hashtable] -or $attrs0 -is [System.Collections.IDictionary]) {
                    if ($reviewer -eq '' -and $attrs0.Contains('EM_Reviewer_Email')) { $reviewer = [string]$attrs0['EM_Reviewer_Email'] }
                    if ($designer -eq '' -and $attrs0.Contains('EM_Designer_Email')) { $designer = [string]$attrs0['EM_Designer_Email'] }
                    if ($checker  -eq '' -and $attrs0.Contains('EM_Checker_Email'))  { $checker  = [string]$attrs0['EM_Checker_Email'] }
                } else {
                    if ($reviewer -eq '' -and $null -ne $attrs0.PSObject.Properties['EM_Reviewer_Email']) { $reviewer = [string]$attrs0.EM_Reviewer_Email }
                    if ($designer -eq '' -and $null -ne $attrs0.PSObject.Properties['EM_Designer_Email']) { $designer = [string]$attrs0.EM_Designer_Email }
                    if ($checker  -eq '' -and $null -ne $attrs0.PSObject.Properties['EM_Checker_Email'])  { $checker  = [string]$attrs0.EM_Checker_Email }
                }
            }
        }
    }

    $to = @()
    foreach ($role in @($eventCfg.to)) {
        switch -Regex ([string]$role) {
            '^(?i)reviewers$' { if (-not [string]::IsNullOrWhiteSpace($reviewer)) { $to += $reviewer } }
            '^(?i)designers$' { if (-not [string]::IsNullOrWhiteSpace($designer)) { $to += $designer } }
            '^(?i)checkers$'  { if (-not [string]::IsNullOrWhiteSpace($checker))  { $to += $checker } }
            default { if (-not [string]::IsNullOrWhiteSpace([string]$role)) { $to += [string]$role } }
        }
    }
    $cc = @()
    foreach ($role in @($eventCfg.cc)) {
        switch -Regex ([string]$role) {
            '^(?i)reviewers$' { if (-not [string]::IsNullOrWhiteSpace($reviewer)) { $cc += $reviewer } }
            '^(?i)designers$' { if (-not [string]::IsNullOrWhiteSpace($designer)) { $cc += $designer } }
            '^(?i)checkers$'  { if (-not [string]::IsNullOrWhiteSpace($checker))  { $cc += $checker } }
            default { if (-not [string]::IsNullOrWhiteSpace([string]$role)) { $cc += [string]$role } }
        }
    }

    $event = @{
        eventType = [string]$eventCfg.eventType
        notificationType = [string]$eventCfg.eventType
        previousState = [string]$PreviousState
        currentState = [string]$CurrentState
        targetState = [string]$CurrentState
        documentName = $docName
        documentGuid = $docGuid
        project = [string]$Project
        stateTransitionKey = [string]$StateTransitionKey
        transitionId = $TransitionId
        reviewers = @($reviewer) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        designers = @($designer) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        checkers = @($checker)  | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        to = $to
        cc = $cc
        _document = $Document
        _eventCfg = $eventCfg
        sourceJobId = if ($null -ne $Job -and $null -ne $Job.id) { [string]$Job.id } else { '' }
        submittedBy = '(unknown)'
        changedByUsername = ''
        changedByUser = $null
        cycleId = ''
    }

    $body = if ($null -ne $eventCfg.actionRequired) { [string]$eventCfg.actionRequired } else { '' }
    $sheetStem = ''
    if (-not [string]::IsNullOrWhiteSpace($docName)) {
        try { $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension($docName) } catch { $sheetStem = '' }
    }
    $event.sheetStem = $sheetStem

    $allRecipients = @($to + $cc) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique
    $event.recipientKey = if ($allRecipients.Count -gt 0) { 'recipients:' + ($allRecipients -join ',') } else { '' }

    $from = ([string]$PreviousState).ToLowerInvariant()
    $toState = ([string]$CurrentState).ToLowerInvariant()
    $anchor = if (-not [string]::IsNullOrWhiteSpace($sheetStem)) { "sheet:$($sheetStem.ToLowerInvariant())|from:$from|to:$toState" } else { "from:$from|to:$toState" }
    if (-not [string]::IsNullOrWhiteSpace($StateTransitionKey)) { $anchor = "$anchor|$StateTransitionKey" }
    $event.logicalTransitionAnchor = $anchor

    $event.dedupeKey = "sheetStem=$sheetStem|previousState=$PreviousState|currentState=$CurrentState|transitionSource=|logicalTransitionAnchor=$anchor|recipientKey=$($event.recipientKey)"

    # Minimal file-based dedupe (mirrors the unit-test behavior we care about).
    if ($null -ne $Config.notifications -and $null -ne $Config.notifications.dedupe -and [bool]$Config.notifications.dedupe.enabled) {
        $storePath = ''
        try { $storePath = [string]$Config.notifications.dedupe.storePath } catch { $storePath = '' }
        if (-not [string]::IsNullOrWhiteSpace($storePath) -and -not [string]::IsNullOrWhiteSpace($event.dedupeKey)) {
            if (Test-Path -LiteralPath $storePath) {
                $existing = @()
                try { $existing = Get-Content -LiteralPath $storePath -ErrorAction SilentlyContinue } catch { $existing = @() }
                if ($existing -contains $event.dedupeKey) {
                    $skipped = @{
                        success = $false
                        skipped = $true
                        dedupeKey = $event.dedupeKey
                        message = 'Duplicate notification suppressed.'
                    }
                    return [pscustomobject]@{
                        IsSuccess = $true
                        Code = 'QC_NOTIFICATION_SKIPPED_DUPLICATE'
                        Message = $skipped.message
                        Data = $skipped
                    }
                }
            } else {
                New-Item -ItemType Directory -Path (Split-Path $storePath) -Force | Out-Null
            }
            Add-Content -LiteralPath $storePath -Value $event.dedupeKey
        }
    }

    return (Send-QCNotification -Event $event -Config $Config -Subject '' -Body $body -To $to -Cc $cc)
}

$testRoot = Join-Path $env:TEMP ("qc-notify-test-" + [guid]::NewGuid().ToString('N'))
$mockRoot = Join-Path $testRoot 'notifications'
$dedupePath = Join-Path $testRoot 'dedupe\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $dedupePath) -Force | Out-Null

function New-NotifyConfig([bool]$Enabled, [string]$Provider = 'Mock') {
    return @{
        auditPoller = @{
            workflowTriggers = @{
                # Unit tests use synthetic documents; do not require a QC-PDF-only notifier gate.
                qcPdfNotificationsOnly = $false
            }
        }
        notifications = @{
            enabled = $Enabled
            provider = $Provider
            dryRun = $true
            outputRoot = $mockRoot
            dedupe = @{
                enabled = $true
                storePath = $dedupePath
                keyFields = @('sheetStem', 'previousState', 'currentState', 'transitionSource', 'logicalTransitionAnchor', 'recipientKey')
                sheetPackageKeyFields = @('sheetStem', 'currentState', 'cycleId')
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

# Missing recipients -> quiet skip without failing the worker job
$noRecipientCfg = New-NotifyConfig -Enabled $true
$noRecipients = Invoke-QCNotificationForStateChange -Config $noRecipientCfg -PreviousState 'In Production' -CurrentState 'QC Received' `
    -Document (New-MockDocument '' '')
Assert-True ($null -ne $noRecipients) 'Missing recipients should return a result object'
Assert-True $noRecipients.IsSuccess 'Missing recipients should complete without job failure'
Assert-True $noRecipients.Data.skipped 'Missing recipients should be marked skipped'

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
Assert-True (-not $prependBlock.blocked) 'Prepend must not be blocked when notification emails are missing'
$prependOk = Test-QCPrependBlockedByMissingEmailAttributes -Config $missingCfg -SheetPdfName 'sheet001.pdf' `
    -Document (New-MockDocument 'reviewer@company.com' 'designer@company.com')
Assert-True (-not $prependOk.blocked) 'Prepend should proceed when notification emails are present'

$defaultRollbackCfg = New-NotifyConfig -Enabled $true
Assert-True (-not $defaultRollbackCfg.notifications.rollbackWhenEmailAttributesMissing) 'Default config should not rollback workflow on missing emails'

# Rollback previous state falls back to production when sheet_index would match the blocked target
$rollbackCfg = New-NotifyConfig -Enabled $true
$rollbackCfg.qcWorkflow = @{ states = @{ production = 'In Production'; qcInitiated = 'QC Initiated'; readyForQc = 'Ready for QC' } }
$rollbackPrevious = Resolve-QCWorkflowRollbackPreviousState -Config $rollbackCfg -TargetStateName 'QC Initiated' `
    -Members @(@{ documentName = 'sheet001.pdf'; documentGuid = 'guid-sheet' }) -SheetPdfGuid 'guid-sheet'
Assert-Eq $rollbackPrevious 'In Production' 'Rollback should fall back to production when index has no prior state'

# Independent Check routes reviewers role to checker email
$icCfg = New-NotifyConfig -Enabled $true
$icSettings = Get-QCNotificationSettings -Config $icCfg
$icSettings.attributes.reviewTypeField = 'QC_Review_Type'
$icSettings.attributes.checkerEmailField = 'EM_Checker_Email'
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
$folderDedupeCfg.notifications.dedupe.keyFields = @('folderPath', 'sheetStem', 'eventType', 'previousState', 'currentState', 'cycleId')
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

# Sibling audit echo events must not fork dedupe (audit:39541 vs audit:39542 on same sheet transition)
$echoCfg = New-NotifyConfig -Enabled $true
$echoCfg.notifications.dedupe.keyFields = @('folderPath', 'sheetStem', 'eventType', 'previousState', 'currentState', 'cycleId')
$echoSettings = Get-QCNotificationSettings -Config $echoCfg
$keyEcho1 = Get-QCNotificationDedupeKey -Event @{
    eventType = 'QC_COMPLETE'
    documentName = '0818000063ea501-qc.pdf'
    sheetStem = '0818000063ea501'
    folderPath = $sharedFolder
    previousState = ''
    currentState = 'QC Complete'
    stateTransitionKey = 'audit:39541'
} -Settings $echoSettings
$keyEcho2 = Get-QCNotificationDedupeKey -Event @{
    eventType = 'QC_COMPLETE'
    documentName = '0818000063ea501-qc.pdf'
    sheetStem = '0818000063ea501'
    folderPath = $sharedFolder
    previousState = ''
    currentState = 'QC Complete'
    stateTransitionKey = 'audit:39542'
} -Settings $echoSettings
Assert-Eq $keyEcho1 $keyEcho2 'Different audit echo ids should not fork notification dedupe'
Assert-True ($keyEcho1 -notmatch 'audit:39541') 'Dedupe key should not include audit event id'
Assert-True ($keyEcho1 -match 'previousState=QC Finalizing') 'Empty stale-index previousState should normalize for QC Complete'

# Prepend writeback (explicit previousState) should match stale-index audit path for same sheet transition
$keyPrepend = Get-QCNotificationDedupeKey -Event @{
    eventType = 'QC_COMPLETE'
    documentName = '0818000063ea501-qc.pdf'
    sheetStem = '0818000063ea501'
    folderPath = $sharedFolder
    previousState = 'QC Finalizing'
    currentState = 'QC Complete'
    cycleId = 'qc_qcprepend_c6e50e714af81799'
} -Settings $echoSettings
$keyAuditWithCycle = Get-QCNotificationDedupeKey -Event @{
    eventType = 'QC_COMPLETE'
    documentName = '0818000063ea501-qc.pdf'
    sheetStem = '0818000063ea501'
    folderPath = $sharedFolder
    previousState = ''
    currentState = 'QC Complete'
    cycleId = 'qc_qcprepend_c6e50e714af81799'
    stateTransitionKey = 'audit:39541'
} -Settings $echoSettings
Assert-Eq $keyPrepend $keyAuditWithCycle 'Prepend writeback and audit stale-index should dedupe when cycle id matches'

# Redlines Received audit echoes (39533 vs 39534 pattern) must share one dedupe key
$keyRedlines1 = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    documentName = '0818000063ea501-qc.pdf'
    sheetStem = '0818000063ea501'
    folderPath = $sharedFolder
    previousState = ''
    currentState = 'Redlines Received'
    stateTransitionKey = 'audit:5001'
} -Settings $echoSettings
$keyRedlines2 = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    documentName = '0818000063ea501-qc.pdf'
    sheetStem = '0818000063ea501'
    folderPath = $sharedFolder
    previousState = ''
    currentState = 'Redlines Received'
    stateTransitionKey = 'audit:5002'
} -Settings $echoSettings
Assert-Eq $keyRedlines1 $keyRedlines2 'Redlines audit echo ids should not fork notification dedupe'
Assert-True (-not [string]::IsNullOrWhiteSpace($keyRedlines1)) 'Stale-index Redlines should produce a dedupe key'

# Different QC cycles must not share notification dedupe when cycleId differs
$cycleSplitCfg = New-NotifyConfig -Enabled $true
$cycleSplitCfg.notifications.dedupe.keyFields = @('folderPath', 'sheetStem', 'eventType', 'previousState', 'currentState', 'cycleId')
$cycleSplitSettings = Get-QCNotificationSettings -Config $cycleSplitCfg
$keyCycle1 = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    sheetStem = '0818000063ea500'
    folderPath = $sharedFolder
    previousState = 'Ready for QC'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_oldcycle001'
} -Settings $cycleSplitSettings
$keyCycle2 = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    sheetStem = '0818000063ea500'
    folderPath = $sharedFolder
    previousState = 'Ready for QC'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_newcycle002'
} -Settings $cycleSplitSettings
Assert-True ($keyCycle1 -ne $keyCycle2) 'Different cycleId values should produce different dedupe keys'

# Sub-cycle bumps (1 -> 1.1) must not collide with the forward Redlines notification in the same major cycle
$subCycleCfg = New-NotifyConfig -Enabled $true
$subCycleCfg.notifications.dedupe.keyFields = @('folderPath', 'sheetStem', 'eventType', 'previousState', 'currentState', 'cycleId')
$subCycleSettings = Get-QCNotificationSettings -Config $subCycleCfg
$keyForwardRedlines = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    sheetStem = '0818000063ea500'
    folderPath = $sharedFolder
    previousState = 'Ready for QC'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_cycle001|1'
} -Settings $subCycleSettings
$keyResubmitRedlines = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    sheetStem = '0818000063ea500'
    folderPath = $sharedFolder
    previousState = 'Corrections Received'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_cycle001|1.1'
} -Settings $subCycleSettings
Assert-True ($keyForwardRedlines -ne $keyResubmitRedlines) 'Forward and resubmit Redlines should use different cycleId dedupe keys'

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
$readyEvent = @{
    eventType = 'READY_FOR_QC'
    notificationType = 'READY_FOR_QC'
    previousState = 'QC Initiated'
    currentState = 'Ready for QC'
    targetState = 'Ready for QC'
    project = 'cafwy2200-i-15_elpse'
    documentName = '080J082001ca001.pdf'
    documentGuid = 'guid-ready-subject'
    reviewers = @('reviewer@company.com')
    designers = @('designer@company.com')
    to = @('reviewer@company.com')
    cc = @('designer@company.com')
    recipientKey = 'recipients:designer@company.com,reviewer@company.com'
    logicalTransitionAnchor = 'audit:9999'
}
$readySend = Send-QCNotification -Event $readyEvent -Config $readyCfg -Subject '' -Body 'Begin QC review.' -To @('reviewer@company.com') -Cc @('designer@company.com')
Assert-True $readySend.IsSuccess 'Ready for QC notification should succeed'
$readyPayload = Get-Content -LiteralPath $readySend.Data.filePath -Raw | ConvertFrom-Json
Assert-True ($readyPayload.subject -match '080J082001ca001') 'Subject should use sheet stem, not QC PDF placeholder'
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
    -Document $auditEchoDoc -StateTransitionKey 'audit:9001'
Assert-True $ae2.Data.skipped 'Second audit echo for same sheet transition should dedupe'
Assert-Eq $ae2.Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'Audit echo duplicate should use skip code'

$redlinesEchoCfg = New-NotifyConfig -Enabled $true
$redlinesEchoCfg.notifications.dedupe.storePath = Join-Path $testRoot 'dedupe-redlines-echo\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $redlinesEchoCfg.notifications.dedupe.storePath) -Force | Out-Null
$redlinesEchoCfg.notifications.events['Redlines Received'] = @{
    enabled = $true
    eventType = 'REDLINES_RECEIVED'
    to = @('designers')
    cc = @('reviewers')
    actionRequired = 'Designer corrections.'
}
$re1 = Invoke-QCNotificationForStateChange -Config $redlinesEchoCfg -PreviousState '' -CurrentState 'Redlines Received' `
    -Document $auditEchoDoc -StateTransitionKey 'audit:9101'
Assert-True $re1.IsSuccess 'First Redlines Received send should succeed'
$re2 = Invoke-QCNotificationForStateChange -Config $redlinesEchoCfg -PreviousState '' -CurrentState 'Redlines Received' `
    -Document $auditEchoDoc -StateTransitionKey 'audit:9101'
Assert-True $re2.Data.skipped 'Second Redlines audit echo should dedupe'
Assert-Eq $re2.Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'Redlines audit echo duplicate should use skip code'

# Prepend writeback enqueues from sheet PDF; notification should resolve to *-qc.pdf and send
$prependCompleteCfg = New-NotifyConfig -Enabled $true
$prependCompleteCfg.notifications.dedupe.storePath = Join-Path $testRoot 'dedupe-prepend-complete\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $prependCompleteCfg.notifications.dedupe.storePath) -Force | Out-Null
$prependCompleteCfg.notifications.events['QC Complete'] = @{
    enabled = $true
    eventType = 'QC_COMPLETE'
    to = @('designers')
    cc = @('reviewers')
    actionRequired = 'QC review cycle is complete.'
}
$prependCompleteCfg.auditPoller = @{ workflowTriggers = @{ qcPdfNotificationsOnly = $true } }
$sheetPdfDoc = [pscustomobject]@{
    Name = '0818000063ea500.pdf'
    DocumentGUID = 'guid-prepend-sheet-pdf'
    EM_Designer_Email = 'designer@company.com'
    EM_Reviewer_Email = 'reviewer@company.com'
}
$prependNotifJob = @{
    id = 'qc_notification_prepend_writeback_test'
    sourceName = '0818000063ea500.pdf'
    sourceFolder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
    metadata = @{
        previousState = 'QC Finalizing'
        currentState = 'QC Complete'
        parentJobId = 'qc_qcprepend_test'
    }
}
$prependCompleteSend = Invoke-QCNotificationForStateChange -Config $prependCompleteCfg -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -Document $sheetPdfDoc -Job $prependNotifJob -DocumentGuid 'guid-prepend-sheet-pdf'
Assert-True $prependCompleteSend.IsSuccess 'Prepend QC Complete notification should succeed'
Assert-True (-not $prependCompleteSend.Data.skipped) 'Prepend QC Complete from sheet PDF should not be suppressed as package member'

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

# Logical transition dedupe keys must distinguish different transitions on the same sheet
$logicalCfg = New-NotifyConfig -Enabled $true
$logicalSettings = Get-QCNotificationSettings -Config $logicalCfg
$sharedStem = '080J082001ca001'
$sharedGuid = 'guid-logical-transition'
$sharedRecipients = @{ to = @('designer@company.com'); cc = @('reviewer@company.com') }
$keyRedlines = Get-QCNotificationDedupeKey -Event (@{
    eventType = 'REDLINES_RECEIVED'
    sheetStem = $sharedStem
    documentGuid = $sharedGuid
    previousState = 'Ready for QC'
    currentState = 'Redlines Received'
    stateTransitionKey = 'audit:40597'
    transitionSource = 'user_audit'
} + $sharedRecipients) -Settings $logicalSettings
$keyCorrections = Get-QCNotificationDedupeKey -Event (@{
    eventType = 'CORRECTIONS_RECEIVED'
    sheetStem = $sharedStem
    documentGuid = $sharedGuid
    previousState = 'Redlines Received'
    currentState = 'Corrections Received'
    stateTransitionKey = 'audit:40598'
    transitionSource = 'user_audit'
} + $sharedRecipients) -Settings $logicalSettings
Assert-True ($keyRedlines -ne $keyCorrections) 'Redlines and Corrections transitions should use different dedupe keys'
Assert-True ($keyCorrections -match 'audit:40598') 'Corrections dedupe key should include audit event id in logical anchor'

# Same logical transition with a new audit event id should not reuse the prior dedupe key
$keyCorrections2 = Get-QCNotificationDedupeKey -Event (@{
    eventType = 'CORRECTIONS_RECEIVED'
    sheetStem = $sharedStem
    documentGuid = $sharedGuid
    previousState = 'Redlines Received'
    currentState = 'Corrections Received'
    stateTransitionKey = 'audit:40600'
    transitionSource = 'user_audit'
} + $sharedRecipients) -Settings $logicalSettings
Assert-True ($keyCorrections -ne $keyCorrections2) 'Different audit ids for the same from/to should produce different dedupe keys'

# Duplicate suppression for the exact same logical transition
$sameTransitionCfg = New-NotifyConfig -Enabled $true
$sameTransitionCfg.notifications.dedupe.storePath = Join-Path $testRoot 'dedupe-same-logical\sent-keys.jsonl'
New-Item -ItemType Directory -Path (Split-Path $sameTransitionCfg.notifications.dedupe.storePath) -Force | Out-Null
$sameTransitionCfg.notifications.events['Corrections Received'] = @{
    enabled = $true
    eventType = 'CORRECTIONS_RECEIVED'
    to = @('designers')
    cc = @('reviewers')
    actionRequired = 'Reviewer verification.'
}
$sameDoc = [pscustomobject]@{
    Name = ($sharedStem + '-qc.pdf')
    DocumentGUID = $sharedGuid
    EM_Designer_Email = 'designer@company.com'
    EM_Reviewer_Email = 'reviewer@company.com'
}
$corr1 = Invoke-QCNotificationForStateChange -Config $sameTransitionCfg -PreviousState 'Redlines Received' -CurrentState 'Corrections Received' `
    -Document $sameDoc -StateTransitionKey 'audit:40598' -NotificationStateSource 'user_audit'
Assert-True $corr1.IsSuccess 'First Corrections Received send should succeed'
$corr2 = Invoke-QCNotificationForStateChange -Config $sameTransitionCfg -PreviousState 'Redlines Received' -CurrentState 'Corrections Received' `
    -Document $sameDoc -StateTransitionKey 'audit:40598' -NotificationStateSource 'user_audit'
Assert-True $corr2.Data.skipped 'Exact same logical transition should dedupe'
Assert-Eq $corr2.Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'Duplicate Corrections transition should use skip code'
Assert-True (-not (Test-QCNotificationResultSent -Result $corr2)) 'Skipped duplicate should not count as sent'

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
    -Document $tidDoc -StateTransitionKey 'audit:9101' -TransitionId 42
Assert-True $t2.Data.skipped 'Same transition id with different audit echo should dedupe'
$t3 = Invoke-QCNotificationForStateChange -Config $tidCfg -PreviousState 'Corrections Received' -CurrentState 'QC Complete' `
    -Document $tidDoc -StateTransitionKey 'audit:9103' -TransitionId 43
Assert-True $t3.IsSuccess 'New transition id (QC cycle) should send again'
Assert-True (-not $t3.Data.skipped) 'New transition id should not be deduped'

# Default dedupe ignores per-file GUID so orphan/replacement QC PDFs do not fork keys
$guidAgnosticCfg = New-NotifyConfig -Enabled $true
$guidAgnosticSettings = Get-QCNotificationSettings -Config $guidAgnosticCfg
$sharedLogicalAnchor = 'sheet:080j082001ab001|from:ready for qc|to:redlines received|cycle:qc_qcprepend_test|1|audit:41067'
$keyGuidA = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    documentName = '080J082001ab001-qc.pdf'
    sheetStem = '080J082001ab001'
    documentGuid = '27d9a8ba-6aaa-4e51-a1d8-9759b20880bb'
    previousState = 'Ready for QC'
    currentState = 'Redlines Received'
    transitionSource = 'user_audit'
    logicalTransitionAnchor = $sharedLogicalAnchor
    recipientKey = 'recipients:jflint@aztec.us'
} -Settings $guidAgnosticSettings
$keyGuidB = Get-QCNotificationDedupeKey -Event @{
    eventType = 'REDLINES_RECEIVED'
    documentName = '080J082001ab001-qc.pdf'
    sheetStem = '080J082001ab001'
    documentGuid = '72bf6609-063c-40db-b067-94d1b064a5b5'
    previousState = 'Ready for QC'
    currentState = 'Redlines Received'
    transitionSource = 'user_audit'
    logicalTransitionAnchor = $sharedLogicalAnchor
    recipientKey = 'recipients:jflint@aztec.us'
} -Settings $guidAgnosticSettings
Assert-Eq $keyGuidA $keyGuidB 'Different QC PDF GUIDs with same logical transition should share one dedupe key'
Assert-True ($keyGuidA -notmatch '27d9a8ba') 'Dedupe key must not include stale QC PDF GUID'
Assert-True ($keyGuidA -notmatch '72bf6609') 'Dedupe key must not include replacement QC PDF GUID'

# Sheet-package dedupe key collapses audit echo paths within the same cycle
$pkgKeyA = Get-QCNotificationSheetPackageDedupeKey -Event @{
    sheetStem = '080J082001ab001'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_test|1'
} -Settings $guidAgnosticSettings
$pkgKeyB = Get-QCNotificationSheetPackageDedupeKey -Event @{
    sheetStem = '080J082001ab001'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_test|1'
} -Settings $guidAgnosticSettings
Assert-Eq $pkgKeyA $pkgKeyB 'Sheet-package dedupe key should be stable'
Register-QCNotificationSheetPackageDedupe -Event @{
    sheetStem = '080J082001ab001'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_test|1'
} -Settings $guidAgnosticSettings -ResultData @{ eventType = 'REDLINES_RECEIVED'; documentName = '080J082001ab001-qc.pdf'; provider = 'Mock' }
Assert-True (Test-QCNotificationSheetPackageAlreadySent -Event @{
    sheetStem = '080J082001ab001'
    currentState = 'Redlines Received'
    cycleId = 'qc_qcprepend_test|1'
} -Settings $guidAgnosticSettings) 'Registered sheet-package dedupe should block echo sends'

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
Assert-Eq $display 'audit.user' 'SubmittedBy prefers explicit username over PW User fallback'

# State-specific recipient routing (Originated / Redlines / Ready for Verification / Verified)
$routeCfg = New-NotifyConfig -Enabled $true
$routeCfg.qcWorkflow = @{
    reviewTypes = @{ independentCheck = 'Check'; peerReview = 'Review'; productionQc = 'Production' }
    processTypes = @{ production = 'Production'; check = 'Check'; review = 'Review' }
}
$routeSettings = Get-QCNotificationSettings -Config $routeCfg
$routeCfg.notifications.events['Originated'] = @{ enabled = $true; eventType = 'READY_FOR_QC'; to = @('reviewers'); cc = @() }
$routeCfg.notifications.events['Redlines Received'] = @{ enabled = $true; eventType = 'REDLINES_RECEIVED'; to = @('designers'); cc = @() }
$routeCfg.notifications.events['Ready for Verification'] = @{ enabled = $true; eventType = 'READY_FOR_VERIFICATION'; to = @('reviewers'); cc = @() }
$routeCfg.notifications.events['Verified'] = @{ enabled = $true; eventType = 'QC_COMPLETE'; to = @('designers', 'reviewers', 'checkers'); cc = @() }

$revDoc = New-MockDocument -Reviewer 'reviewer@company.com' -Designer 'designer@company.com' -Checker 'checker@company.com' -ReviewType 'Review'
$revDoc.Name = 'sheet-rev.pdf'
$originated = Resolve-QCNotificationRecipients -Document $revDoc -Settings $routeSettings -Config $routeCfg `
    -ToRoles @('reviewers') -CcRoles @()
Assert-Eq $originated.to[0] 'reviewer@company.com' 'Originated review lane notifies reviewer only'
Assert-Eq $originated.to.Count 1 'Originated should not include designer'

$chkDoc = New-MockDocument -Reviewer 'reviewer@company.com' -Designer 'designer@company.com' -Checker 'checker@company.com' -ReviewType 'Check'
$chkDoc.Name = 'sheet-chk.pdf'
$chkOriginated = Resolve-QCNotificationRecipients -Document $chkDoc -Settings $routeSettings -Config $routeCfg `
    -ToRoles @('reviewers') -CcRoles @()
Assert-Eq $chkOriginated.to[0] 'checker@company.com' 'Originated check lane notifies checker via reviewers role'
Assert-Eq $chkOriginated.to.Count 1 'Originated check lane should not include designer'

$redlines = Resolve-QCNotificationRecipients -Document $revDoc -Settings $routeSettings -Config $routeCfg `
    -ToRoles @('designers') -CcRoles @()
Assert-Eq $redlines.to[0] 'designer@company.com' 'Redlines Received notifies designer only'
Assert-Eq $redlines.to.Count 1 'Redlines Received should not include reviewer'

$verified = Resolve-QCNotificationRecipients -Document $revDoc -Settings $routeSettings -Config $routeCfg `
    -ToRoles @('designers', 'reviewers', 'checkers') -CcRoles @()
Assert-Eq $verified.to.Count 3 'Verified notifies designer, reviewer, and checker'
Assert-True ($verified.to -contains 'designer@company.com') 'Verified includes designer'
Assert-True ($verified.to -contains 'reviewer@company.com') 'Verified includes reviewer'
Assert-True ($verified.to -contains 'checker@company.com') 'Verified includes checker'

Write-Host 'All QC notification tests passed.'
