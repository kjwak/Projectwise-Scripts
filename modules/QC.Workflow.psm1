# QC.Workflow.psm1
# Responsibility: Configurable ProjectWise QC workflow/state/attribute writeback framework.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'Core.Logging.psm1') -Force -ErrorAction SilentlyContinue
if (-not (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'QC.Notifications.psm1') -Force -ErrorAction SilentlyContinue
}
Import-Module (Join-Path $PSScriptRoot 'QC.AuditTriggers.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue

function _QCW-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return @{ value = $Value } }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCW-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCW-StateNameEquals([string]$Left, [string]$Right) {
    if (_QCW-IsNullOrWhiteSpace $Left) { return _QCW-IsNullOrWhiteSpace $Right }
    if (_QCW-IsNullOrWhiteSpace $Right) { return $false }
    return ([string]$Left).Trim().Equals(([string]$Right).Trim(), [System.StringComparison]::OrdinalIgnoreCase)
}

function _QCW-GetConfiguredWorkflowStateLabels([hashtable]$Settings) {
    $labels = [System.Collections.Generic.List[string]]::new()
    if (-not $Settings) { return @() }
    $states = _QCW-ToHashtable $Settings.states
    if ($states) {
        foreach ($k in $states.Keys) {
            $v = [string]$states[$k]
            if (-not (_QCW-IsNullOrWhiteSpace $v)) {
                $t = $v.Trim()
                if (-not $labels.Contains($t)) { $labels.Add($t) | Out-Null }
            }
        }
    }
    foreach ($k in @('stateAfterSuccessfulPrepend', 'stateAfterFailedPrepend', 'defaultStateAfterPrepend')) {
        if ($Settings.ContainsKey($k) -and -not (_QCW-IsNullOrWhiteSpace $Settings[$k])) {
            $t = ([string]$Settings[$k]).Trim()
            if (-not $labels.Contains($t)) { $labels.Add($t) | Out-Null }
        }
    }
    $byTrigger = _QCW-ToHashtable $Settings.stateAfterPrependByTrigger
    if ($byTrigger) {
        foreach ($k in $byTrigger.Keys) {
            $v = [string]$byTrigger[$k]
            if (-not (_QCW-IsNullOrWhiteSpace $v)) {
                $t = $v.Trim()
                if (-not $labels.Contains($t)) { $labels.Add($t) | Out-Null }
            }
        }
    }
    return @($labels)
}

function _QCW-TestTargetInConfiguredStates([hashtable]$Settings, [string]$TargetStateName) {
    if (_QCW-IsNullOrWhiteSpace $TargetStateName) { return $false }
    foreach ($label in @(_QCW-GetConfiguredWorkflowStateLabels -Settings $Settings)) {
        if (_QCW-StateNameEquals $label $TargetStateName) { return $true }
    }
    return $false
}

function _QCW-GetPropertyValue([object]$Object, [string[]]$Names) {
    foreach ($name in @($Names)) {
        try {
            if ($Object -and $Object.PSObject -and $Object.PSObject.Properties[$name] -and $null -ne $Object.$name) {
                return $Object.$name
            }
        } catch { }
    }
    return $null
}

function _QCW-Log([string]$Event, [string]$Level, [string]$Message, [hashtable]$Data) {
    try {
        $payload = @{}
        if ($Data) { foreach ($k in $Data.Keys) { $payload[$k] = $Data[$k] } }
        if (Get-Command -Name Write-QCJsonLog -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level $Level -Code $Event -Message $Message -Data $payload | Out-Null
            return
        }
        if (Get-Command -Name Write-QCLog -ErrorAction SilentlyContinue) {
            $payload['event'] = $Event
            Write-QCLog -Level $Level -Message $Message -Data $payload | Out-Null
        }
    } catch { }
}

function _QCW-NewWorkflowResult([bool]$IsSuccess, [string]$Code, [string]$Message, [hashtable]$Data) {
    if ($IsSuccess) { return New-QCSuccessResult -Code $Code -Message $Message -Data $Data }
    return New-QCFailureResult -Code $Code -Message $Message -Data $Data
}

function Invoke-QCWorkflowStateChangeNotification {
    <#
    .SYNOPSIS
    Enqueues or sends a workflow state-change notification via the shared prepend/worker path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [hashtable]$Context,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [object]$Document = $null
    )

    if ($Context -and (Get-Command -Name 'Advance-QCWorkflowCycleForRedlinesResubmit' -ErrorAction SilentlyContinue)) {
        try {
            $wfSettings = Get-QCWorkflowSettings -Config $Config
            $Context = Advance-QCWorkflowCycleForRedlinesResubmit -Settings $wfSettings -Context $Context -Config $Config `
                -PreviousState $PreviousState -CurrentState $CurrentState
        } catch { }
    }

    return _QCW-InvokeStateChangeNotification -Config $Config -Context $Context `
        -PreviousState $PreviousState -CurrentState $CurrentState -Document $Document
}

function _QCW-ParseQCCycleNumber {
    param([object]$Value)

    $major = 0
    $minor = 0
    if ($null -eq $Value) {
        return @{ major = 0; minor = 0; display = '' }
    }
    $text = [string]$Value
    if (_QCW-IsNullOrWhiteSpace $text) {
        return @{ major = 0; minor = 0; display = '' }
    }
    $text = $text.Trim()
    if ($text -match '^(\d+)\.(\d+)$') {
        try {
            $major = [int]$Matches[1]
            $minor = [int]$Matches[2]
        } catch { }
    } elseif ($text -match '^(\d+)$') {
        try {
            $major = [int]$Matches[1]
            $minor = 0
        } catch { }
    }
    $display = if ($minor -gt 0) { '{0}.{1}' -f $major, $minor } elseif ($major -gt 0) { [string]$major } else { '' }
    return @{ major = $major; minor = $minor; display = $display }
}

function _QCW-FormatQCCycleNumber {
    param(
        [int]$Major,
        [int]$Minor = 0
    )

    if ($Major -le 0) { return '' }
    if ($Minor -gt 0) { return '{0}.{1}' -f $Major, $Minor }
    return [string]$Major
}

function _QCW-GetBaseQCCycleId {
    param([string]$CycleId)

    if (_QCW-IsNullOrWhiteSpace $CycleId) { return '' }
    $id = [string]$CycleId.Trim()
    $pipe = $id.IndexOf('|')
    if ($pipe -gt 0) { return $id.Substring(0, $pipe) }
    return $id
}

function _QCW-BuildQCCycleId {
    param(
        [string]$BaseCycleId,
        [string]$CycleNumberDisplay
    )

    $base = _QCW-GetBaseQCCycleId $BaseCycleId
    if (_QCW-IsNullOrWhiteSpace $base) { return '' }
    if (_QCW-IsNullOrWhiteSpace $CycleNumberDisplay) { return $base }
    return '{0}|{1}' -f $base, ([string]$CycleNumberDisplay).Trim()
}

function _QCW-ResolveSheetCycleTargets {
    param([hashtable]$Context)

    $folderPath = ''
    $sheetStem = ''
    $sourceDocGuid = ''
    $job = $null
    if ($Context -and $Context.ContainsKey('job') -and $Context.job) { $job = $Context.job }
    if ($Context -and $Context.ContainsKey('documentGuid') -and $Context.documentGuid) {
        $sourceDocGuid = [string]$Context.documentGuid
    }
    if ($job) {
        if (-not $folderPath -and $job.sourceFolder) { $folderPath = [string]$job.sourceFolder }
        if (-not $sourceDocGuid) {
            $md = _QCW-ToHashtable $job.metadata
            if ($md -and $md.triggerDocumentGuid) { $sourceDocGuid = [string]$md.triggerDocumentGuid }
            elseif ($md -and $md.documentGuid) { $sourceDocGuid = [string]$md.documentGuid }
        }
        if ($job.sourceName) {
            $srcName = [string]$job.sourceName
            if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                try { $sheetStem = [string](Get-PWSheetStemFromDocumentName -DocumentName $srcName) } catch { }
            }
            if (_QCW-IsNullOrWhiteSpace $sheetStem) {
                $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension($srcName)
            }
        }
    }
    if ((_QCW-IsNullOrWhiteSpace $folderPath) -and $Context -and $Context.ContainsKey('folderPath') -and $Context.folderPath) {
        $folderPath = [string]$Context.folderPath
    }
    if ((_QCW-IsNullOrWhiteSpace $folderPath) -and $Context -and $Context.ContainsKey('sourceFolder') -and $Context.sourceFolder) {
        $folderPath = [string]$Context.sourceFolder
    }
    if ((_QCW-IsNullOrWhiteSpace $sheetStem) -and $Context -and $Context.ContainsKey('sourceName') -and $Context.sourceName) {
        $srcName2 = [string]$Context.sourceName
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            try { $sheetStem = [string](Get-PWSheetStemFromDocumentName -DocumentName $srcName2) } catch { }
        }
        if (_QCW-IsNullOrWhiteSpace $sheetStem) {
            $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension($srcName2)
        }
    }
    return @{
        folderPath = $folderPath
        sheetStem = $sheetStem
        documentGuid = $sourceDocGuid
    }
}

function _QCW-ApplyCycleToContext {
    param(
        [hashtable]$Context,
        [string]$CycleId,
        [string]$CycleNumber
    )

    if (-not $Context) { return $Context }
    if (-not $Context.ContainsKey('attributes') -or -not $Context.attributes) {
        $Context['attributes'] = @{}
    }
    $attrs = _QCW-ToHashtable $Context.attributes
    if (-not $attrs) { $attrs = @{} }
    $attrs['cycleId'] = $CycleId
    $attrs['cycleNumber'] = $CycleNumber
    $Context['attributes'] = $attrs
    $Context['cycleId'] = $CycleId
    $Context['cycleNumber'] = $CycleNumber
    return $Context
}

function _QCW-EnsureNotificationCommandsLoaded {
    if (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue) { return $true }
    $path = Join-Path $PSScriptRoot 'QC.Notifications.psm1'
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    try { Import-Module $path -Force -ErrorAction Stop | Out-Null } catch { return $false }
    return [bool](Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)
}

function _QCW-ResolveNotificationRoleAttrsForEnqueue {
    param(
        [hashtable]$Config,
        [object]$Document,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = ''
    )

    $out = @{}
    $attr = @{}
    if ($Config -and (Get-Command -Name 'Get-QCNotificationSettings' -ErrorAction SilentlyContinue)) {
        try {
            $settings = Get-QCNotificationSettings -Config $Config
            $attr = _QCW-ToHashtable $settings.attributes
            if (-not $attr) { $attr = @{} }
        } catch { }
    }
    $reviewerField = if ($attr.reviewerEmailField) { [string]$attr.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $designerField = if ($attr.designerEmailField) { [string]$attr.designerEmailField } else { 'EM_Designer_Email' }
    $checkerField = if ($attr.checkerEmailField) { [string]$attr.checkerEmailField } else { 'EM_Checker_Email' }

    if ($Document) {
        $docH = _QCW-ToHashtable $Document
        if ($docH) {
            if ($docH.ContainsKey('designerEmail') -and -not (_QCW-IsNullOrWhiteSpace $docH.designerEmail)) {
                $out['designerEmail'] = [string]$docH.designerEmail
            } elseif ($docH.ContainsKey($designerField) -and -not (_QCW-IsNullOrWhiteSpace $docH[$designerField])) {
                $out['designerEmail'] = [string]$docH[$designerField]
            }
            if ($docH.ContainsKey('reviewerEmail') -and -not (_QCW-IsNullOrWhiteSpace $docH.reviewerEmail)) {
                $out['reviewerEmail'] = [string]$docH.reviewerEmail
            } elseif ($docH.ContainsKey($reviewerField) -and -not (_QCW-IsNullOrWhiteSpace $docH[$reviewerField])) {
                $out['reviewerEmail'] = [string]$docH[$reviewerField]
            }
            if ($docH.ContainsKey('checkerEmail') -and -not (_QCW-IsNullOrWhiteSpace $docH.checkerEmail)) {
                $out['checkerEmail'] = [string]$docH.checkerEmail
            } elseif ($docH.ContainsKey($checkerField) -and -not (_QCW-IsNullOrWhiteSpace $docH[$checkerField])) {
                $out['checkerEmail'] = [string]$docH[$checkerField]
            }
        }
    }

    $srcName = [string]$SourceDocumentName
    if ((_QCW-IsNullOrWhiteSpace $srcName) -and $Document) {
        try { $srcName = [string]$Document.Name } catch { }
    }
    if (-not (_QCW-IsNullOrWhiteSpace $srcName) -and $srcName -match '(?i)-qc\.pdf$') {
        $srcName = [string]([System.IO.Path]::GetFileNameWithoutExtension($srcName)) + '.pdf'
    }
    $folder = [string]$FolderPath
    if ((_QCW-IsNullOrWhiteSpace $folder) -and $Document) {
        try { $folder = [string]$Document.FolderPath } catch { }
    }

    $needPw = (-not $out.ContainsKey('reviewerEmail')) -or (-not $out.ContainsKey('designerEmail')) -or (-not $out.ContainsKey('checkerEmail'))
    if ($needPw -and $Config -and -not (_QCW-IsNullOrWhiteSpace $folder) -and -not (_QCW-IsNullOrWhiteSpace $srcName)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            try {
                $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $folder -SourceDocumentName $srcName -Config $Config
                if ($pw.found) {
                    if ((-not $out.ContainsKey('reviewerEmail')) -and $pw.reviewerEmail) { $out['reviewerEmail'] = [string]$pw.reviewerEmail }
                    if ((-not $out.ContainsKey('designerEmail')) -and $pw.designerEmail) { $out['designerEmail'] = [string]$pw.designerEmail }
                    if ((-not $out.ContainsKey('checkerEmail')) -and $pw.checkerEmail) { $out['checkerEmail'] = [string]$pw.checkerEmail }
                    if ($pw.qcReviewType) { $out['qcReviewType'] = [string]$pw.qcReviewType }
                }
            } catch { }
        }
    }
    return $out
}

function _QCW-InvokeStateChangeNotification {
    param(
        [hashtable]$Config,
        [hashtable]$Context,
        [string]$PreviousState,
        [string]$CurrentState,
        [object]$Document
    )

    if (-not $Config) { return $null }
    if (-not (_QCW-EnsureNotificationCommandsLoaded)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_NOTIFICATION_MODULE_UNAVAILABLE' `
                -Message 'QC.Notifications module not loaded; state-change notification skipped.' -Data @{
                currentState = [string]$CurrentState
                previousState = [string]$PreviousState
            } | Out-Null
        }
        return New-QCFailureResult -Code 'QC_NOTIFICATION_MODULE_UNAVAILABLE' -Message 'QC.Notifications module not loaded.' -Data @{}
    }

    $ctxChangedByUser = $null
    $ctxChangedByUsername = ''
    if ($Context) {
        if ($Context.ContainsKey('changedByUser') -and $null -ne $Context.changedByUser) {
            try {
                $n = [int]$Context.changedByUser
                if ($n -gt 0) { $ctxChangedByUser = $n }
            } catch { }
        }
        if ($Context.ContainsKey('changedByUsername') -and -not (_QCW-IsNullOrWhiteSpace $Context.changedByUsername)) {
            $ctxChangedByUsername = [string]$Context.changedByUsername
        }
    }

    $job = $null
    if ($Context -and $Context.ContainsKey('job')) { $job = $Context.job }

    $enqueueAsJob = $false
    if (Get-Command -Name 'Test-QCNotificationsEnqueueAsJob' -ErrorAction SilentlyContinue) {
        try { $enqueueAsJob = Test-QCNotificationsEnqueueAsJob -Config $Config } catch { $enqueueAsJob = $false }
    }
    if ($enqueueAsJob -and (Get-Command -Name 'Add-QCQueueJob' -ErrorAction SilentlyContinue)) {
        $dedupeKey = $null
        if (Get-Command -Name 'Get-QCNotificationDedupeKey' -ErrorAction SilentlyContinue) {
            try {
                $notifSettings = Get-QCNotificationSettings -Config $Config
                $docGuid = ''
                $docName = ''
                if ($Document) {
                    try { $docGuid = [string]$Document.DocumentGUID } catch { }
                    try { if (-not $docGuid) { $docGuid = [string]$Document.DocumentGuid } } catch { }
                    try { $docName = [string]$Document.Name } catch { }
                }
                $eventTypeForDedupe = ([string]$CurrentState).ToUpperInvariant().Replace(' ', '_')
                $eventsMap = _QCW-ToHashtable $notifSettings.events
                if ($eventsMap -and $eventsMap.ContainsKey([string]$CurrentState)) {
                    $evCfg = _QCW-ToHashtable $eventsMap[[string]$CurrentState]
                    if ($evCfg -and $evCfg.eventType) { $eventTypeForDedupe = [string]$evCfg.eventType }
                }
                $eventForDedupe = @{
                    eventType = $eventTypeForDedupe
                    documentGuid = $docGuid
                    documentName = $docName
                    previousState = [string]$PreviousState
                    currentState = [string]$CurrentState
                }
                if ($Context -and $Context.folderPath) {
                    $eventForDedupe['folderPath'] = [string]$Context.folderPath
                } elseif ($job -and $job.sourceFolder) {
                    $eventForDedupe['folderPath'] = [string]$job.sourceFolder
                } elseif ($Context -and $Context.documentPath -and ([string]$Context.documentPath -match '\\')) {
                    $eventForDedupe['folderPath'] = [System.IO.Path]::GetDirectoryName([string]$Context.documentPath)
                    $eventForDedupe['documentPath'] = [string]$Context.documentPath
                }
                if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue -and -not (_QCW-IsNullOrWhiteSpace $docName)) {
                    try { $eventForDedupe['sheetStem'] = [string](Get-PWSheetStemFromDocumentName -DocumentName $docName) } catch { }
                }
                if ($Context -and $Context.ContainsKey('transitionId') -and $null -ne $Context.transitionId) {
                    try {
                        $ctxTid = [int]$Context.transitionId
                        if ($ctxTid -gt 0) { $eventForDedupe['transitionId'] = $ctxTid }
                    } catch { }
                }
                $stKeyForDedupe = $null
                if ($Context -and $Context.ContainsKey('stateTransitionKey') -and $Context.stateTransitionKey) {
                    $stKeyForDedupe = [string]$Context.stateTransitionKey
                } elseif ($job -and $job.metadata -is [hashtable] -and $job.metadata.ContainsKey('stateTransitionKey') -and $job.metadata.stateTransitionKey) {
                    $stKeyForDedupe = [string]$job.metadata.stateTransitionKey
                }
                if (-not (_QCW-IsNullOrWhiteSpace $stKeyForDedupe)) { $eventForDedupe['stateTransitionKey'] = $stKeyForDedupe }
                if ($job -and $job.metadata -is [hashtable]) {
                    $wfMdForDedupe = _QCW-ToHashtable $job.metadata
                    if ($wfMdForDedupe -and $wfMdForDedupe.attributes) {
                        $eventForDedupe['attributes'] = _QCW-ToHashtable $wfMdForDedupe.attributes
                    }
                    if ($wfMdForDedupe -and $wfMdForDedupe.cycleId -and -not $eventForDedupe.ContainsKey('cycleId')) {
                        $eventForDedupe['cycleId'] = [string]$wfMdForDedupe.cycleId
                    }
                }
                if ($Context -and $Context.ContainsKey('attributes') -and $Context.attributes) {
                    $ctxAttrs = _QCW-ToHashtable $Context.attributes
                    if ($ctxAttrs) {
                        if ($ctxAttrs.ContainsKey('cycleId') -and -not (_QCW-IsNullOrWhiteSpace $ctxAttrs.cycleId)) {
                            $eventForDedupe['cycleId'] = [string]$ctxAttrs.cycleId
                        }
                        $eventForDedupe['attributes'] = $ctxAttrs
                    }
                } elseif ($Context -and $Context.ContainsKey('cycleId') -and -not (_QCW-IsNullOrWhiteSpace $Context.cycleId)) {
                    $eventForDedupe['cycleId'] = [string]$Context.cycleId
                }
                $dedupeKey = Get-QCNotificationDedupeKey -Event $eventForDedupe -Settings $notifSettings -Config $Config -Job $job
            } catch { }
        }
        if ($dedupeKey) {
            $dup = Test-QCDuplicateJob -DedupeKey $dedupeKey -Config $Config
            if ($dup.IsSuccess -and [bool]$dup.Data.isDuplicate) {
                $pendingDup = $false
                foreach ($match in @($dup.Data.matches)) {
                    if (-not $match) { continue }
                    $matchState = ''
                    try { $matchState = [string]$match.state } catch { }
                    if ($matchState -in @('pending', 'running')) { $pendingDup = $true; break }
                }
                if ($pendingDup) {
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Level 'Information' -Code 'QC_NOTIFICATION_ENQUEUE_SKIPPED_DUPLICATE' `
                            -Message 'Notification job already queued for this dedupe key.' -Data @{
                            dedupeKey = $dedupeKey
                            currentState = [string]$CurrentState
                            previousState = [string]$PreviousState
                        } | Out-Null
                    }
                    return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' -Message 'Notification job already queued.' -Data @{ dedupeKey = $dedupeKey }
                }
            }
            if (Get-Command -Name 'Test-QCNotificationDedupe' -ErrorAction SilentlyContinue) {
                try {
                    $notifSettingsForSent = Get-QCNotificationSettings -Config $Config
                    $ctxTransitionId = $null
                    if ($Context -and $Context.ContainsKey('transitionId') -and $null -ne $Context.transitionId) {
                        try { $ctxTransitionId = [int]$Context.transitionId } catch { }
                    }
                    if (Test-QCNotificationDedupe -DedupeKey $dedupeKey -Settings $notifSettingsForSent -Config $Config -TransitionId $ctxTransitionId) {
                        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                            Write-QCJsonLog -Level 'Information' -Code 'QC_NOTIFICATION_ENQUEUE_SKIPPED_ALREADY_SENT' `
                                -Message 'Notification already sent for this logical sheet transition.' -Data @{
                                dedupeKey = $dedupeKey
                                currentState = [string]$CurrentState
                                previousState = [string]$PreviousState
                            } | Out-Null
                        }
                        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' -Message 'Notification already sent for this sheet transition.' -Data @{ dedupeKey = $dedupeKey }
                    }
                } catch { }
            }
        }
        $wfTransitionKey = $null
        if ($Context -and $Context.ContainsKey('stateTransitionKey') -and $Context.stateTransitionKey) {
            $wfTransitionKey = [string]$Context.stateTransitionKey
        } elseif ($job -and $job.metadata -is [hashtable] -and $job.metadata.ContainsKey('stateTransitionKey') -and $job.metadata.stateTransitionKey) {
            $wfTransitionKey = [string]$job.metadata.stateTransitionKey
        } elseif ($job -and $job.id) {
            $wfTransitionKey = 'workflow:job:' + [string]$job.id + '|state:' + ([string]$CurrentState).Trim()
        }
        $wfChangedByUser = $null
        $wfChangedByUsername = ''
        if ($job -and $job.ContainsKey('metadata') -and $job.metadata) {
            $wfMd = _QCW-ToHashtable $job.metadata
            if ($wfMd) {
                if ($wfMd.ContainsKey('changedByUser') -and $null -ne $wfMd.changedByUser) {
                    try {
                        $n = [int]$wfMd.changedByUser
                        if ($n -gt 0) { $wfChangedByUser = $n }
                    } catch { }
                }
                if ($wfMd.ContainsKey('changedByUsername') -and -not (_QCW-IsNullOrWhiteSpace $wfMd.changedByUsername)) {
                    $wfChangedByUsername = [string]$wfMd.changedByUsername
                }
            }
        }
        $notifJob = @{
            id = ('qc_notification_' + [guid]::NewGuid().ToString('n').Substring(0, 12))
            type = 'QC_NOTIFICATION'
            status = 'pending'
            createdAt = (Get-Date).ToUniversalTime().ToString('o')
            attempts = 0
            sourceFolder = if ($job -and $job.sourceFolder) { [string]$job.sourceFolder } else { '' }
            sourceName = if ($Document) { try { [string]$Document.Name } catch { '' } } else { '' }
            dedupeKey = $dedupeKey
            metadata = @{
                previousState = $PreviousState
                currentState = $CurrentState
                documentGuid = if ($Document) { try { [string]$Document.DocumentGUID } catch { '' } } else { '' }
                parentJobId = if ($Context -and $Context.jobId) { [string]$Context.jobId } elseif ($job -and $job.id) { [string]$job.id } else { '' }
                stateTransitionKey = $wfTransitionKey
            }
        }
        if ($null -ne $wfChangedByUser) { $notifJob.metadata['changedByUser'] = $wfChangedByUser }
        if (-not (_QCW-IsNullOrWhiteSpace $wfChangedByUsername)) { $notifJob.metadata['changedByUsername'] = $wfChangedByUsername }
        if ($null -eq $wfChangedByUser -and $null -ne $ctxChangedByUser) { $notifJob.metadata['changedByUser'] = $ctxChangedByUser }
        if ((_QCW-IsNullOrWhiteSpace $wfChangedByUsername) -and -not (_QCW-IsNullOrWhiteSpace $ctxChangedByUsername)) {
            $notifJob.metadata['changedByUsername'] = $ctxChangedByUsername
        }
        if ($Context -and $Context.ContainsKey('transitionId') -and $null -ne $Context.transitionId) {
            try {
                $ctxTid = [int]$Context.transitionId
                if ($ctxTid -gt 0) { $notifJob.metadata['transitionId'] = $ctxTid }
            } catch { }
        }
        if ($Context -and $Context.ContainsKey('documentPath') -and $Context.documentPath) {
            $dp = [string]$Context.documentPath
            if ($dp -match '\\') { $notifJob.sourceFolder = [System.IO.Path]::GetDirectoryName($dp) }
            if (-not $notifJob.sourceName) { $notifJob.sourceName = [System.IO.Path]::GetFileName($dp) }
        }
        if (-not $notifJob.sourceFolder -and $Context -and $Context.folderPath) {
            $notifJob.sourceFolder = [string]$Context.folderPath
        }
        if ($Context -and $Context.ContainsKey('sourceName') -and -not (_QCW-IsNullOrWhiteSpace $Context.sourceName)) {
            $notifJob.sourceName = [string]$Context.sourceName
        } elseif (-not $notifJob.sourceName -and $Document) {
            try {
                $derivedName = [string]$Document.Name
                if ($derivedName -match '(?i)-qc\.pdf$') {
                    $notifJob.sourceName = [System.IO.Path]::GetFileNameWithoutExtension($derivedName) + '.pdf'
                } elseif (-not (_QCW-IsNullOrWhiteSpace $derivedName)) {
                    $notifJob.sourceName = $derivedName
                }
            } catch { }
        }
        if ($job) {
            if (-not $notifJob.sourceFolder -and $job.sourceFolder) { $notifJob.sourceFolder = [string]$job.sourceFolder }
            if (-not $notifJob.sourceName -and $job.sourceName) { $notifJob.sourceName = [string]$job.sourceName }
        }
        $roleAttrs = $null
        if ($Context -and $Context.ContainsKey('attributes') -and $Context.attributes) {
            $roleAttrs = _QCW-ToHashtable $Context.attributes
        }
        if (-not $roleAttrs -and $job -and $job.ContainsKey('metadata') -and $job.metadata) {
            $jobMd = _QCW-ToHashtable $job.metadata
            if ($jobMd -and $jobMd.attributes) { $roleAttrs = _QCW-ToHashtable $jobMd.attributes }
        }
        $resolvedRoles = _QCW-ResolveNotificationRoleAttrsForEnqueue -Config $Config -Document $Document `
            -FolderPath ([string]$notifJob.sourceFolder) -SourceDocumentName ([string]$notifJob.sourceName)
        if ($resolvedRoles -and $resolvedRoles.Count -gt 0) {
            if (-not $roleAttrs) { $roleAttrs = @{} }
            foreach ($k in @('designerEmail', 'reviewerEmail', 'checkerEmail', 'reviewType', 'qcReviewType')) {
                if ($resolvedRoles.ContainsKey($k) -and -not (_QCW-IsNullOrWhiteSpace $resolvedRoles[$k])) {
                    if (-not $roleAttrs.ContainsKey($k) -or (_QCW-IsNullOrWhiteSpace $roleAttrs[$k])) {
                        $roleAttrs[$k] = [string]$resolvedRoles[$k]
                    }
                }
            }
        }
        if ($roleAttrs) {
            $attrsForNotif = @{}
            foreach ($k in @('designerEmail', 'reviewerEmail', 'checkerEmail', 'reviewType', 'qcReviewType', 'cycleId', 'cycleNumber')) {
                if ($roleAttrs.ContainsKey($k) -and -not (_QCW-IsNullOrWhiteSpace $roleAttrs[$k])) {
                    $attrsForNotif[$k] = [string]$roleAttrs[$k]
                }
            }
            if ($attrsForNotif.Count -gt 0) {
                $notifJob.metadata['attributes'] = $attrsForNotif
            }
        }
        if ($Context -and $Context.ContainsKey('cycleId') -and -not (_QCW-IsNullOrWhiteSpace $Context.cycleId)) {
            $notifJob.metadata['cycleId'] = [string]$Context.cycleId
            if (-not $notifJob.metadata.ContainsKey('attributes') -or -not $notifJob.metadata.attributes) {
                $notifJob.metadata['attributes'] = @{}
            }
            $notifJob.metadata['attributes']['cycleId'] = [string]$Context.cycleId
            if ($Context.ContainsKey('cycleNumber') -and $null -ne $Context.cycleNumber) {
                $notifJob.metadata['attributes']['cycleNumber'] = [string]$Context.cycleNumber
            }
        }
        if (Get-Command -Name '_QCN-ResolveQcPdfNotificationTarget' -ErrorAction SilentlyContinue) {
            try {
                $docGuidForResolve = if ($notifJob.metadata.documentGuid) { [string]$notifJob.metadata.documentGuid } else { '' }
                $qcEnqueueTarget = _QCN-ResolveQcPdfNotificationTarget -Document $Document -Config $Config -Job $notifJob `
                    -DocumentName ([string]$notifJob.sourceName) -DocumentGuid $docGuidForResolve `
                    -DocumentPath ([string]$notifJob.sourcePath)
                if ($qcEnqueueTarget -and ([string]$qcEnqueueTarget.documentName -match '(?i)-qc\.pdf$')) {
                    $notifJob.sourceName = [string]$qcEnqueueTarget.documentName
                    if (-not (_QCW-IsNullOrWhiteSpace $qcEnqueueTarget.documentGuid)) {
                        $notifJob.metadata['documentGuid'] = [string]$qcEnqueueTarget.documentGuid
                    }
                }
            } catch { }
        }
        if (-not (_QCW-IsNullOrWhiteSpace $notifJob.sourceFolder) -and -not (_QCW-IsNullOrWhiteSpace $notifJob.sourceName)) {
            $notifJob.sourcePath = Join-Path ([string]$notifJob.sourceFolder) ([string]$notifJob.sourceName)
        }
        if (_QCW-IsNullOrWhiteSpace $dedupeKey) {
            $fallbackParts = @(
                ('currentState={0}' -f ([string]$CurrentState).Trim())
                ('previousState={0}' -f ([string]$PreviousState).Trim())
            )
            if ($Context -and $Context.folderPath) { $fallbackParts += ('folderPath={0}' -f [string]$Context.folderPath) }
            if ($Context -and $Context.ContainsKey('stateTransitionKey') -and $Context.stateTransitionKey) {
                $fallbackParts += ('stateTransitionKey={0}' -f [string]$Context.stateTransitionKey)
            } elseif (-not (_QCW-IsNullOrWhiteSpace $wfTransitionKey)) {
                $fallbackParts += ('stateTransitionKey={0}' -f $wfTransitionKey)
            }
            if (-not (_QCW-IsNullOrWhiteSpace $notifJob.sourceName)) {
                $fallbackParts += ('sourceName={0}' -f [string]$notifJob.sourceName)
            }
            $dedupeKey = ($fallbackParts -join '|')
            $notifJob.dedupeKey = $dedupeKey
        }
        $enq = Add-QCQueueJob -Job $notifJob -Config $Config
        if ($enq.IsSuccess) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'QC_NOTIFICATION_ENQUEUED' `
                    -Message 'Notification deferred to QC_NOTIFICATION job.' -Data @{
                    jobId = [string]$notifJob.id
                    dedupeKey = $dedupeKey
                    sourceFolder = [string]$notifJob.sourceFolder
                    sourceName = [string]$notifJob.sourceName
                    previousState = [string]$PreviousState
                    currentState = [string]$CurrentState
                } | Out-Null
            }
            return New-QCSuccessResult -Code 'QC_NOTIFICATION_ENQUEUED' -Message 'Notification deferred to QC_NOTIFICATION job.' -Data @{ jobId = [string]$notifJob.id }
        }
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_NOTIFICATION_ENQUEUE_FAILED' `
                -Message 'Failed to enqueue QC_NOTIFICATION job; falling back to inline send.' -Data @{
                jobId = [string]$notifJob.id
                dedupeKey = $dedupeKey
                enqueueCode = if ($enq -and $enq.Code) { [string]$enq.Code } else { '' }
                enqueueMessage = if ($enq -and $enq.Message) { [string]$enq.Message } else { '' }
                currentState = [string]$CurrentState
                previousState = [string]$PreviousState
            } | Out-Null
        }
    }
    try {
        $notifJob = $job
        if ($Context -and $Context.ContainsKey('attributes') -and $Context.attributes) {
            if (-not $notifJob) { $notifJob = @{} }
            if (-not $notifJob.ContainsKey('metadata') -or -not $notifJob.metadata) { $notifJob['metadata'] = @{} }
            $notifJob.metadata['attributes'] = $Context.attributes
        }
        if ($Context -and $Context.ContainsKey('jobId') -and $notifJob -and -not $notifJob.ContainsKey('id')) {
            $notifJob['id'] = [string]$Context.jobId
        }
        if ($Context -and $Context.ContainsKey('documentPath') -and $notifJob) {
            if (-not $notifJob.ContainsKey('sourceFolder')) {
                $dp = [string]$Context.documentPath
                if ($dp -match '\\') { $notifJob['sourceFolder'] = [System.IO.Path]::GetDirectoryName($dp) }
            }
            if (-not $notifJob.ContainsKey('sourceName') -and $Document) {
                try { $notifJob['sourceName'] = [string]$Document.Name } catch { }
            }
        }
        $stKey = $null
        if ($Context -and $Context.ContainsKey('stateTransitionKey') -and $Context.stateTransitionKey) {
            $stKey = [string]$Context.stateTransitionKey
        } elseif ($job -and $job.metadata -is [hashtable] -and $job.metadata.ContainsKey('stateTransitionKey') -and $job.metadata.stateTransitionKey) {
            $stKey = [string]$job.metadata.stateTransitionKey
        } elseif ($job -and $job.id) {
            $stKey = 'workflow:job:' + [string]$job.id + '|state:' + ([string]$CurrentState).Trim()
        }
        $wfChangedByUser = $null
        $wfChangedByUsername = ''
        if ($job -and $job.ContainsKey('metadata') -and $job.metadata) {
            $wfMd = _QCW-ToHashtable $job.metadata
            if ($wfMd) {
                if ($wfMd.ContainsKey('changedByUser') -and $null -ne $wfMd.changedByUser) {
                    try {
                        $n = [int]$wfMd.changedByUser
                        if ($n -gt 0) { $wfChangedByUser = $n }
                    } catch { }
                }
                if ($wfMd.ContainsKey('changedByUsername') -and -not (_QCW-IsNullOrWhiteSpace $wfMd.changedByUsername)) {
                    $wfChangedByUsername = [string]$wfMd.changedByUsername
                }
            }
        }
        if (($null -eq $wfChangedByUser) -and (_QCW-IsNullOrWhiteSpace $wfChangedByUsername) -and $null -ne $ctxChangedByUser) {
            $wfChangedByUser = $ctxChangedByUser
        }
        if ((_QCW-IsNullOrWhiteSpace $wfChangedByUsername) -and -not (_QCW-IsNullOrWhiteSpace $ctxChangedByUsername)) {
            $wfChangedByUsername = $ctxChangedByUsername
        }
        if (($null -eq $wfChangedByUser) -and (_QCW-IsNullOrWhiteSpace $wfChangedByUsername) -and $Config) {
            $ap = _QCW-ToHashtable $Config.auditPoller
            if ($ap) {
                $wt = _QCW-ToHashtable $ap.workflowTriggers
                if ($wt -and $wt.automationPwUsernames) {
                    foreach ($name in @($wt.automationPwUsernames)) {
                        if (-not (_QCW-IsNullOrWhiteSpace $name)) { $wfChangedByUsername = [string]$name; break }
                    }
                }
            }
        }
        $notifyParams = @{
            Config = $Config
            PreviousState = $PreviousState
            CurrentState = $CurrentState
            Document = $Document
            Job = $notifJob
            StateTransitionKey = $stKey
            ChangedByUser = $wfChangedByUser
            ChangedByUsername = $wfChangedByUsername
        }
        if ($Context -and $Context.ContainsKey('transitionId') -and $null -ne $Context.transitionId) {
            try {
                $ctxTid = [int]$Context.transitionId
                if ($ctxTid -gt 0) { $notifyParams['TransitionId'] = $ctxTid }
            } catch { }
        }
        return Invoke-QCNotificationForStateChange @notifyParams
    } catch {
        if (Get-Command -Name Write-QCNotificationResult -ErrorAction SilentlyContinue) {
            Write-QCNotificationResult -Code 'QC_NOTIFICATION_HOOK_FAILED' -Level 'Error' -Message $_.Exception.Message `
                -Result @{ success = $false } -Event @{ previousState = $PreviousState; currentState = $CurrentState }
        }
        return $null
    }
}

function _QCW-GetCommandParameterName([string]$CommandName, [string[]]$CandidateNames) {
    $cmd = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }
    foreach ($name in @($CandidateNames)) {
        if ($cmd.Parameters.ContainsKey($name)) { return $name }
    }
    return $null
}

function _QCW-InvokeGetPWWorkflowStateLinks([string]$WorkflowName) {
    <#
    Calls Get-PWWorkflowStateLinks with an explicit workflow name only.
    Never invokes the cmdlet without WorkflowName (avoids interactive prompts on pwps_dab).
    #>
    if (_QCW-IsNullOrWhiteSpace $WorkflowName) { return @() }

    $cmd = Get-Command -Name 'Get-PWWorkflowStateLinks' -ErrorAction SilentlyContinue
    if (-not $cmd) { return @() }

    $name = $WorkflowName.Trim()
    $attempts = [System.Collections.Generic.List[hashtable]]::new()
    $wfParam = _QCW-GetCommandParameterName -CommandName 'Get-PWWorkflowStateLinks' -CandidateNames @('WorkflowName', 'Workflow')
    if ($wfParam) {
        $attempts.Add(@{ $wfParam = $name }) | Out-Null
    } else {
        # pwps_dab often omits parameter metadata; still bind by common names (never call with zero args).
        $attempts.Add(@{ WorkflowName = $name }) | Out-Null
        $attempts.Add(@{ Workflow = $name }) | Out-Null
    }

    $lastError = $null
    foreach ($args in $attempts) {
        try {
            $links = @(& $cmd @args -ErrorAction Stop)
            if ($links.Count -gt 0) { return $links }
        } catch {
            $lastError = $_
        }
    }

    try {
        return @(& $cmd $name -ErrorAction Stop)
    } catch {
        if ($lastError) { throw $lastError }
        throw
    }
}

function _QCW-ObjectNameMatches([object]$Object, [string]$ExpectedName, [string[]]$PropertyNames) {
    if (_QCW-IsNullOrWhiteSpace $ExpectedName) { return $false }
    foreach ($p in @($PropertyNames)) {
        $v = _QCW-GetPropertyValue -Object $Object -Names @($p)
        if (-not (_QCW-IsNullOrWhiteSpace $v) -and ([string]$v).Trim() -eq $ExpectedName) { return $true }
    }
    if ($Object -is [string] -and $Object.Trim() -eq $ExpectedName) { return $true }
    return $false
}

function _QCW-DefaultWorkflowStates {
    return @{
        production = 'In Production'
        qcInitiated = 'QC Initiated'
        qcReceived = 'Ready for QC'
        readyForQc = 'Ready for QC'
        redlinesReceived = 'Redlines Received'
        correctionsReceived = 'Corrections Received'
        qcFinalizing = 'QC Finalizing'
        complete = 'QC Complete'
        error = 'Error Needs Attention'
    }
}

function _QCW-DefaultReviewTypes {
    return @{
        productionQc = 'Production QC'
        peerReview = 'Peer Review'
        independentCheck = 'Independent Check'
    }
}

function _QCW-DefaultAttributeMap {
    return @{
        qcActive = 'QC_Active'
        reviewType = 'QC_Review_Type'
        cycleId = 'QC_Cycle_ID'
        cycleNumber = 'QC_Cycle_Number'
        designerEmail = 'QC_Designer_Email'
        reviewerEmail = 'QC_Reviewer_Email'
        checkerEmail = 'QC_Checker_Email'
        assignedTo = 'QC_Assigned_To'
        lastActionBy = 'QC_Last_Action_By'
        lastActionDate = 'QC_Last_Action_Date'
        status = 'QC_Status'
        historyPdfPath = 'QC_History_PDF_Path'
        latestOverlayPdfPath = 'QC_Latest_Overlay_PDF_Path'
        sourceDocumentPath = 'QC_Source_Document_Path'
        automationLastRun = 'QC_Automation_Last_Run'
        automationResult = 'QC_Automation_Result'
        automationError = 'QC_Automation_Error'
    }
}

function Get-QCWorkflowAttributeWritebackExcludeDefaults {
    <#
    .SYNOPSIS
    Internal keys excluded from ProjectWise attribute writeback (Phase 1: DB/telemetry holds these).
    #>
    [CmdletBinding()]
    param()
    return @(
        'qcActive'
        'lastActionBy'
        'lastActionDate'
        'historyPdfPath'
        'latestOverlayPdfPath'
        'sourceDocumentPath'
        'automationLastRun'
        'automationResult'
        'automationError'
    )
}

function _QCW-ResolveAttributeWritebackExcludeSet {
    param([hashtable]$Settings)

    if ($Settings -and $Settings.ContainsKey('attributeWritebackExcludeDisabled')) {
        try {
            if ([bool]$Settings.attributeWritebackExcludeDisabled) { return @() }
        } catch { }
    }

    $exclude = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($key in @(Get-QCWorkflowAttributeWritebackExcludeDefaults)) {
        if (-not [string]::IsNullOrWhiteSpace($key)) { [void]$exclude.Add([string]$key) }
    }
    if ($Settings -and $Settings.ContainsKey('attributeWritebackExclude') -and $Settings.attributeWritebackExclude) {
        foreach ($key in @($Settings.attributeWritebackExclude)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$key)) { [void]$exclude.Add(([string]$key).Trim()) }
        }
    }
    return @($exclude)
}

function _QCW-FilterAttributeWritebackValues {
    param(
        [hashtable]$Values,
        [hashtable]$Settings
    )

    $filtered = @{}
    $skipped = [System.Collections.Generic.List[string]]::new()
    if (-not $Values) { return @{ values = $filtered; skippedKeys = @() } }

    $exclude = _QCW-ResolveAttributeWritebackExcludeSet -Settings $Settings
    foreach ($key in @($Values.Keys)) {
        $k = [string]$key
        if ($exclude -contains $k) {
            [void]$skipped.Add($k)
            continue
        }
        $filtered[$k] = $Values[$k]
    }
    return @{ values = $filtered; skippedKeys = @($skipped) }
}

function Get-QCWorkflowDeprecationWarnings {
    [CmdletBinding()]
    param([hashtable]$RawWorkflowConfig)

    $warnings = [System.Collections.Generic.List[string]]::new()
    if (-not $RawWorkflowConfig) { return @($warnings) }

    $deprecatedKeys = @(
        'productionStateName'
        'receivedStateName'
        'correctionsInProgressStateName'
        'backcheckInProgressStateName'
        'errorStateName'
        'stageMap'
    )
    $deprecatedStateKeys = @('reviewInProgress', 'correctionsInProgress', 'verificationInProgress', 'redlinesIssued')
    foreach ($key in $deprecatedKeys) {
        if ($RawWorkflowConfig.ContainsKey($key)) {
            $warnings.Add("qcWorkflow.$key is deprecated; use qcWorkflow.states and qcWorkflow.reviewTypes instead.") | Out-Null
        }
    }

    $attrMap = _QCW-ToHashtable $RawWorkflowConfig.attributeMap
    if ($attrMap -and ($attrMap.ContainsKey('stage') -or $attrMap.ContainsKey('reviewer'))) {
        $warnings.Add('qcWorkflow.attributeMap keys stage/reviewer are deprecated; lifecycle is ProjectWise state and QC_Review_Type role emails.') | Out-Null
    }
    if ($RawWorkflowConfig.ContainsKey('stageMap')) {
        $warnings.Add('qcWorkflow.stageMap (red/green/blue) is deprecated and ignored; ProjectWise document state is the lifecycle source of truth.') | Out-Null
    }

    $rawStates = _QCW-ToHashtable $RawWorkflowConfig.states
    if ($rawStates) {
        foreach ($sk in $deprecatedStateKeys) {
            if ($rawStates.ContainsKey($sk)) {
                $warnings.Add("qcWorkflow.states.$sk is deprecated; use ownership handoff states (redlinesReceived, correctionsReceived, qcFinalizing).") | Out-Null
            }
        }
    }
    $rawTriggers = _QCW-ToHashtable $RawWorkflowConfig.stateAfterPrependByTrigger
    if ($rawTriggers) {
        foreach ($tk in @('reviewerRedlineUpdate', 'designerCorrectionComplete')) {
            if ($rawTriggers.ContainsKey($tk)) {
                $warnings.Add("qcWorkflow.stateAfterPrependByTrigger.$tk is deprecated; prepend triggers are initialQcPdf and finalQcComplete only.") | Out-Null
            }
        }
    }

    return @($warnings)
}

function _QCW-DefaultPrependStateTriggers {
    $states = _QCW-DefaultWorkflowStates
    return @{
        initialQcPdf = [string]$states.readyForQc
        finalQcComplete = [string]$states.complete
    }
}

function Normalize-QCPrependTriggerKey {
    [CmdletBinding()]
    param([string]$Value)

    if (_QCW-IsNullOrWhiteSpace $Value) { return $null }
    $v = ([string]$Value).Trim()
    $compact = ($v.ToLowerInvariant() -replace '[^a-z0-9]', '')
    switch ($compact) {
        'initialqcpdf' { return 'initialQcPdf' }
        'initial' { return 'initialQcPdf' }
        'qcarchivist' { return 'initialQcPdf' }
        'readyforqc' { return 'initialQcPdf' }
        'qcinitiated' { return 'initialQcPdf' }
        'qcreceived' { return 'initialQcPdf' }
        'finalqccomplete' { return 'finalQcComplete' }
        'qcfinalizing' { return 'finalQcComplete' }
        'finalprepend' { return 'finalQcComplete' }
        default {
            if ($v -eq 'initialQcPdf' -or $v -eq 'finalQcComplete') { return $v }
            return $v
        }
    }
}

function Resolve-QCWorkflowStateAfterPrepend {
    <#
    .SYNOPSIS
    Resolve target ProjectWise state after successful QC_PREPEND from trigger/context (not a single fixed state).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings,
        [hashtable]$Context
    )

    if ($Context -and $Context.ContainsKey('targetState') -and -not (_QCW-IsNullOrWhiteSpace $Context.targetState)) {
        return [string]$Context.targetState
    }

    $triggerRaw = $null
    if ($Context) {
        if ($Context.ContainsKey('prependTrigger') -and $Context.prependTrigger) { $triggerRaw = [string]$Context.prependTrigger }
        elseif ($Context.ContainsKey('job') -and $Context.job) {
            $job = _QCW-ToHashtable $Context.job
            if ($job -and $job.ContainsKey('metadata') -and $job.metadata) {
                $md = _QCW-ToHashtable $job.metadata
                if ($md) {
                    foreach ($k in @('prependTrigger','qcPrependTrigger','workflowPrependTrigger')) {
                        if ($md.ContainsKey($k) -and $md[$k]) { $triggerRaw = [string]$md[$k]; break }
                    }
                }
            }
        }
    }

    $triggerKey = Normalize-QCPrependTriggerKey -Value $triggerRaw
    $map = _QCW-ToHashtable $Settings.stateAfterPrependByTrigger
    if (-not $map) { $map = _QCW-DefaultPrependStateTriggers }

    if ($triggerKey -and $map.ContainsKey($triggerKey) -and -not (_QCW-IsNullOrWhiteSpace $map[$triggerKey])) {
        return [string]$map[$triggerKey]
    }

    if (-not (_QCW-IsNullOrWhiteSpace $Settings.stateAfterSuccessfulPrepend)) {
        return [string]$Settings.stateAfterSuccessfulPrepend
    }
    return [string](_QCW-DefaultWorkflowStates).readyForQc
}

function Get-QCWorkflowStateName {
    [CmdletBinding()]
    param(
        [hashtable]$Settings,
        [Parameter(Mandatory)]
        [string]$StateKey
    )

    $states = _QCW-ToHashtable $Settings.states
    if ($states -and $states.ContainsKey($StateKey) -and -not (_QCW-IsNullOrWhiteSpace $states[$StateKey])) {
        return [string]$states[$StateKey]
    }
    $legacyKeyMap = @{
        redlinesIssued = 'redlinesReceived'
        correctionsInProgress = 'correctionsReceived'
        verificationInProgress = 'correctionsReceived'
        reviewInProgress = 'readyForQc'
    }
    if ($legacyKeyMap.ContainsKey($StateKey) -and $states -and $states.ContainsKey($legacyKeyMap[$StateKey]) -and -not (_QCW-IsNullOrWhiteSpace $states[$legacyKeyMap[$StateKey]])) {
        return [string]$states[$legacyKeyMap[$StateKey]]
    }
    $defaults = _QCW-DefaultWorkflowStates
    if ($defaults.ContainsKey($StateKey)) { return [string]$defaults[$StateKey] }
    if ($legacyKeyMap.ContainsKey($StateKey) -and $defaults.ContainsKey($legacyKeyMap[$StateKey])) {
        return [string]$defaults[$legacyKeyMap[$StateKey]]
    }
    return $null
}

function Resolve-QCWorkflowAssignee {
    <#
    .SYNOPSIS
    Resolve QC_Assigned_To from ProjectWise lifecycle state and QC_Review_Type.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings,
        [string]$StateName,
        [string]$ReviewType,
        [string]$DesignerEmail,
        [string]$ReviewerEmail,
        [string]$CheckerEmail
    )

    $states = _QCW-ToHashtable $Settings.states
    if (-not $states) { $states = _QCW-DefaultWorkflowStates }
    $reviewTypes = _QCW-ToHashtable $Settings.reviewTypes
    if (-not $reviewTypes) { $reviewTypes = _QCW-DefaultReviewTypes }

    $resolvedReviewType = [string]$Settings.defaultReviewType
    if (-not (_QCW-IsNullOrWhiteSpace $ReviewType)) { $resolvedReviewType = [string]$ReviewType }
    if (_QCW-IsNullOrWhiteSpace $resolvedReviewType) { $resolvedReviewType = [string]$reviewTypes.productionQc }

    $independentCheck = [string]$reviewTypes.independentCheck
    $useChecker = (-not (_QCW-IsNullOrWhiteSpace $independentCheck)) -and ($resolvedReviewType.Trim() -eq $independentCheck.Trim())

    $state = if ($StateName) { [string]$StateName } else { '' }
    $production = [string]$states.production
    $qcInitiated = if ($states.qcInitiated) { [string]$states.qcInitiated } else { 'QC Initiated' }
    $ready = if ($states.readyForQc) { [string]$states.readyForQc } elseif ($states.qcReceived) { [string]$states.qcReceived } else { 'Ready for QC' }
    $redlinesReceived = if ($states.redlinesReceived) { [string]$states.redlinesReceived } elseif ($states.redlinesIssued) { [string]$states.redlinesIssued } else { 'Redlines Received' }
    $correctionsReceived = if ($states.correctionsReceived) { [string]$states.correctionsReceived } elseif ($states.verificationInProgress) { [string]$states.verificationInProgress } elseif ($states.correctionsInProgress) { [string]$states.correctionsInProgress } else { 'Corrections Received' }
    $complete = [string]$states.complete

    if ($state -eq $complete) { return $null }

    if ($state -eq $production -or $state -eq $qcInitiated -or $state -eq $redlinesReceived) {
        if (-not (_QCW-IsNullOrWhiteSpace $DesignerEmail)) { return [string]$DesignerEmail }
        return $null
    }
    if ($state -in @($ready, $correctionsReceived)) {
        if ($useChecker) {
            if (-not (_QCW-IsNullOrWhiteSpace $CheckerEmail)) { return [string]$CheckerEmail }
            return $null
        }
        if (-not (_QCW-IsNullOrWhiteSpace $ReviewerEmail)) { return [string]$ReviewerEmail }
        return $null
    }
    return $null
}

function Get-QCWorkflowSettings {
    [CmdletBinding()]
    param([hashtable]$Config)

    $raw = @{}
    if ($Config -and $Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
        $norm = _QCW-ToHashtable $Config.qcWorkflow
        if ($norm) { $raw = $norm }
    }

    $defaultStates = _QCW-DefaultWorkflowStates
    $defaultReviewTypes = _QCW-DefaultReviewTypes

    $settings = @{
        enabled = $false
        strictMode = $false
        dryRunWriteback = $true
        workflowName = ''
        expectedWorkflowName = ''
        mode = 'AttributesOnly'
        states = @{}
        reviewTypes = @{}
        defaultReviewType = 'Production QC'
        defaultStateAfterPrepend = 'Ready for QC'
        stateAfterSuccessfulPrepend = 'Ready for QC'
        stateAfterFailedPrepend = 'Error Needs Attention'
        defaultPrependTrigger = 'initialQcPdf'
        stateAfterPrependByTrigger = @{}
        autoSetState = $false
        autoWriteAttributes = $true
        attributeMap = _QCW-DefaultAttributeMap
        attributeWritebackExcludeDisabled = $false
        attributeWritebackExclude = @()
    }
    foreach ($k in $defaultStates.Keys) { $settings.states[$k] = $defaultStates[$k] }
    foreach ($k in $defaultReviewTypes.Keys) { $settings.reviewTypes[$k] = $defaultReviewTypes[$k] }
    foreach ($k in (_QCW-DefaultPrependStateTriggers).Keys) { $settings.stateAfterPrependByTrigger[$k] = (_QCW-DefaultPrependStateTriggers)[$k] }

    foreach ($k in $raw.Keys) {
        if ($k -eq 'attributeMap') {
            $merged = @{}
            foreach ($ak in $settings.attributeMap.Keys) { $merged[$ak] = $settings.attributeMap[$ak] }
            $incoming = _QCW-ToHashtable $raw.attributeMap
            if ($incoming) {
                foreach ($ak in $incoming.Keys) {
                    if ($ak -eq 'stage') { continue }
                    if ($ak -eq 'reviewer') {
                        if (-not $merged.ContainsKey('reviewerEmail')) { $merged['reviewerEmail'] = $incoming[$ak] }
                        continue
                    }
                    $merged[$ak] = $incoming[$ak]
                }
            }
            $settings.attributeMap = $merged
        }
        elseif ($k -eq 'states' -or $k -eq 'reviewTypes' -or $k -eq 'stateAfterPrependByTrigger') {
            $incoming = _QCW-ToHashtable $raw[$k]
            if ($incoming) {
                foreach ($sk in $incoming.Keys) { $settings[$k][$sk] = $incoming[$sk] }
            }
        }
        elseif ($k -ne 'stageMap') {
            $settings[$k] = $raw[$k]
        }
    }

    # Backward compatibility: legacy flat state name keys override states.* when states.* not in raw config.
    if (-not $raw.ContainsKey('states')) {
        if ($raw.ContainsKey('productionStateName') -and $raw.productionStateName) { $settings.states.production = [string]$raw.productionStateName }
        if ($raw.ContainsKey('receivedStateName') -and $raw.receivedStateName) { $settings.states.qcReceived = [string]$raw.receivedStateName; $settings.states.readyForQc = [string]$raw.receivedStateName }
        if ($raw.ContainsKey('correctionsInProgressStateName') -and $raw.correctionsInProgressStateName) { $settings.states.correctionsReceived = [string]$raw.correctionsInProgressStateName }
        if ($raw.ContainsKey('backcheckInProgressStateName') -and $raw.backcheckInProgressStateName) { $settings.states.correctionsReceived = [string]$raw.backcheckInProgressStateName }
        if ($raw.ContainsKey('errorStateName') -and $raw.errorStateName) { $settings.states.error = [string]$raw.errorStateName }
    }

    # Legacy states.* keys → ownership handoff states
    if ($settings.states.ContainsKey('redlinesIssued') -and -not (_QCW-IsNullOrWhiteSpace $settings.states.redlinesIssued) -and (-not $settings.states.ContainsKey('redlinesReceived') -or (_QCW-IsNullOrWhiteSpace $settings.states.redlinesReceived))) {
        $settings.states.redlinesReceived = [string]$settings.states.redlinesIssued
    }
    if ($settings.states.ContainsKey('verificationInProgress') -and -not (_QCW-IsNullOrWhiteSpace $settings.states.verificationInProgress) -and (-not $settings.states.ContainsKey('correctionsReceived') -or (_QCW-IsNullOrWhiteSpace $settings.states.correctionsReceived))) {
        $settings.states.correctionsReceived = [string]$settings.states.verificationInProgress
    }
    if ($settings.states.ContainsKey('correctionsInProgress') -and -not (_QCW-IsNullOrWhiteSpace $settings.states.correctionsInProgress) -and (-not $settings.states.ContainsKey('correctionsReceived') -or (_QCW-IsNullOrWhiteSpace $settings.states.correctionsReceived))) {
        $settings.states.correctionsReceived = [string]$settings.states.correctionsInProgress
    }
    if ($settings.states.ContainsKey('reviewInProgress') -and -not (_QCW-IsNullOrWhiteSpace $settings.states.reviewInProgress) -and (-not $settings.states.ContainsKey('readyForQc') -or (_QCW-IsNullOrWhiteSpace $settings.states.readyForQc))) {
        $settings.states.readyForQc = [string]$settings.states.reviewInProgress
    }

    if (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName -and -not (_QCW-IsNullOrWhiteSpace $settings.workflowName)) {
        $settings.expectedWorkflowName = $settings.workflowName
    }
    if (_QCW-IsNullOrWhiteSpace $settings.workflowName -and -not (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName)) {
        $settings.workflowName = $settings.expectedWorkflowName
    }
    foreach ($boolKey in @('enabled','strictMode','dryRunWriteback','autoSetState','autoWriteAttributes','attributeWritebackExcludeDisabled')) {
        try { $settings[$boolKey] = [bool]$settings[$boolKey] } catch { }
    }
    if ($settings.attributeWritebackExclude -isnot [System.Array] -and $null -ne $settings.attributeWritebackExclude) {
        $settings.attributeWritebackExclude = @($settings.attributeWritebackExclude)
    } elseif ($null -eq $settings.attributeWritebackExclude) {
        $settings.attributeWritebackExclude = @()
    }
    return $settings
}

function Test-QCWorkflowConfig {
    [CmdletBinding()]
    param([hashtable]$Config)

    $settings = Get-QCWorkflowSettings -Config $Config
    $raw = @{}
    if ($Config -and $Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
        $rawNorm = _QCW-ToHashtable $Config.qcWorkflow
        if ($rawNorm) { $raw = $rawNorm }
    }
    $warnings = [System.Collections.Generic.List[string]]::new()
    $critical = [System.Collections.Generic.List[string]]::new()
    foreach ($w in @(Get-QCWorkflowDeprecationWarnings -RawWorkflowConfig $raw)) { $warnings.Add($w) | Out-Null }

    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'QC_WORKFLOW_CONFIG_DISABLED' -Message 'QC workflow writeback is disabled.' -Data @{ settings = $settings; warnings = @($warnings); criticalIssues = @() }
    }

    $mode = ([string]$settings.mode).Trim()
    if (_QCW-IsNullOrWhiteSpace $mode) { $mode = 'AttributesOnly' }
    if ($mode -notin @('AttributesOnly','StateAndAttributes')) {
        $msg = "qcWorkflow.mode '$mode' is not recognized; use AttributesOnly or StateAndAttributes."
        $warnings.Add($msg) | Out-Null
        $critical.Add($msg) | Out-Null
    }
    if ([bool]$settings.autoSetState -and (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName)) {
        $warnings.Add('qcWorkflow.autoSetState is true but qcWorkflow.expectedWorkflowName is empty; workflow validation will be partial.') | Out-Null
    }
    $rawAttributeMap = $null
    if ($raw.ContainsKey('attributeMap')) { $rawAttributeMap = _QCW-ToHashtable $raw.attributeMap }
    if ([bool]$settings.autoWriteAttributes -and (-not $raw.ContainsKey('attributeMap') -or -not $rawAttributeMap -or $rawAttributeMap.Keys.Count -eq 0)) {
        $msg = 'qcWorkflow.autoWriteAttributes is true but qcWorkflow.attributeMap is missing or empty.'
        $warnings.Add($msg) | Out-Null
        $critical.Add($msg) | Out-Null
    }

    $data = @{ settings = $settings; warnings = @($warnings); criticalIssues = @($critical) }
    if ([bool]$settings.strictMode -and $critical.Count -gt 0) {
        _QCW-Log -Event 'QC_WORKFLOW_STRICT_FAILURE' -Level 'Error' -Message 'QC workflow configuration failed strict validation.' -Data $data
        return New-QCFailureResult -Code 'QC_WORKFLOW_CONFIG_STRICT_FAILURE' -Message 'QC workflow configuration failed strict validation.' -Data $data
    }
    if ($warnings.Count -gt 0) {
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message 'QC workflow configuration warnings were found.' -Data $data
    }
    return New-QCSuccessResult -Code 'QC_WORKFLOW_CONFIG_VALID' -Message 'QC workflow configuration validated.' -Data $data
}

function Get-PWDocumentWorkflowInfo {
    [CmdletBinding()]
    param(
        [object]$Document,
        [hashtable]$Context
    )

    if (-not $Document -and $Context -and $Context.ContainsKey('document')) { $Document = $Context.document }
    $workflowName = _QCW-GetPropertyValue -Object $Document -Names @('WorkflowName','Workflow','WorkflowStateName')
    $stateName = _QCW-GetPropertyValue -Object $Document -Names @('StateName','DocumentState','WorkflowState','State','CurrentState')
    $folderPath = _QCW-GetPropertyValue -Object $Document -Names @('FolderPath','ProjectPath','FullPath')
    if (_QCW-IsNullOrWhiteSpace $folderPath -and $Context -and $Context.ContainsKey('folderPath')) { $folderPath = $Context.folderPath }

    $docGuid = _QCW-GetPropertyValue -Object $Document -Names @('DocumentGUID','DocumentGuid','GUID')
    $docName = _QCW-GetPropertyValue -Object $Document -Names @('Name','FileName','DocumentName')
    if (($Document -or -not (_QCW-IsNullOrWhiteSpace $docGuid) -or (-not (_QCW-IsNullOrWhiteSpace $docName))) -and
        (_QCW-IsNullOrWhiteSpace $stateName) -and (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) {
        try {
            $ctxFolder = if ($Context -and $Context.ContainsKey('sourceFolder')) { [string]$Context.sourceFolder } elseif ($Context -and $Context.job -and $Context.job.sourceFolder) { [string]$Context.job.sourceFolder } else { $null }
            $lookupFolder = if (-not (_QCW-IsNullOrWhiteSpace $folderPath)) { [string]$folderPath } else { $ctxFolder }
            $resolvedState = Get-PWDocumentWorkflowStateName -FolderPath $lookupFolder -DocumentName $docName -DocumentGuid $docGuid
            if (-not (_QCW-IsNullOrWhiteSpace $resolvedState)) { $stateName = $resolvedState }
        } catch { }
    }
    if (_QCW-IsNullOrWhiteSpace $stateName -and $Context) {
        if ($Context.ContainsKey('lifecycleState') -and -not (_QCW-IsNullOrWhiteSpace $Context.lifecycleState)) {
            $stateName = [string]$Context.lifecycleState
        } elseif ($Context.ContainsKey('previousState') -and -not (_QCW-IsNullOrWhiteSpace $Context.previousState)) {
            $stateName = [string]$Context.previousState
        } elseif ($Context.ContainsKey('job') -and $Context.job) {
            $job = _QCW-ToHashtable $Context.job
            if ($job -and $job.ContainsKey('metadata') -and $job.metadata) {
                $md = _QCW-ToHashtable $job.metadata
                if ($md -and $md.ContainsKey('pwStateName') -and $md.pwStateName) { $stateName = [string]$md.pwStateName }
            }
        }
    }

    $workflowNameValue = if (_QCW-IsNullOrWhiteSpace $workflowName) { $null } else { [string]$workflowName }
    $stateNameValue = if (_QCW-IsNullOrWhiteSpace $stateName) { $null } else { [string]$stateName }
    $folderPathValue = if (_QCW-IsNullOrWhiteSpace $folderPath) { $null } else { [string]$folderPath }
    $documentPathValue = if ($Context -and $Context.ContainsKey('documentPath')) { $Context.documentPath } else { $null }

    return New-QCSuccessResult -Code 'QC_WORKFLOW_INFO' -Message 'ProjectWise workflow info resolved from available document properties.' -Data @{
        workflowName = $workflowNameValue
        stateName = $stateNameValue
        folderPath = $folderPathValue
        document = $Document
        documentPath = $documentPathValue
    }
}

function Ensure-PWQCWorkflowAssignment {
    [CmdletBinding()]
    param(
        [hashtable]$Settings,
        [hashtable]$Context,
        [bool]$DryRun
    )

    $document = if ($Context -and $Context.ContainsKey('document')) { $Context.document } else { $null }
    $expected = [string]$Settings.expectedWorkflowName
    $info = Get-PWDocumentWorkflowInfo -Document $document -Context $Context
    $currentWorkflow = $info.Data.workflowName
    $warnings = @()
    $workflowExists = $null

    if (-not (_QCW-IsNullOrWhiteSpace $expected)) {
        $wfCmd = Get-Command -Name Get-PWWorkflows -ErrorAction SilentlyContinue
        if ($wfCmd) {
            try {
                $workflows = @(Get-PWWorkflows -ErrorAction Stop)
                $workflowExists = [bool](@($workflows | Where-Object { _QCW-ObjectNameMatches -Object $_ -ExpectedName $expected -PropertyNames @('Name','WorkflowName','Workflow') }).Count -gt 0)
                if (-not $workflowExists) { $warnings += "Expected ProjectWise workflow '$expected' was not returned by Get-PWWorkflows." }
            } catch {
                $warnings += ('Get-PWWorkflows failed while validating current project workflow: ' + $_.Exception.Message)
            }
        } else {
            $warnings += 'Get-PWWorkflows is unavailable; current project workflow existence cannot be validated.'
        }
    }

    if (_QCW-IsNullOrWhiteSpace $currentWorkflow) {
        $warnings += 'Document object does not expose current workflow; workflow validation is informational only.'
    } elseif (-not (_QCW-IsNullOrWhiteSpace $expected) -and ([string]$currentWorkflow).Trim() -ne $expected.Trim()) {
        $warnings += "Document current workflow '$currentWorkflow' does not match expected workflow '$expected'. QC attributes remain authoritative."
    }

    $data = @{
        expectedWorkflowName = $expected
        currentWorkflowName = $currentWorkflow
        workflowExists = $workflowExists
        documentMatchesExpectedWorkflow = (-not (_QCW-IsNullOrWhiteSpace $currentWorkflow) -and -not (_QCW-IsNullOrWhiteSpace $expected) -and ([string]$currentWorkflow).Trim() -eq $expected.Trim())
        folderPath = $info.Data.folderPath
        dryRun = $DryRun
        changed = $false
        warnings = @($warnings)
        informationalOnly = $true
    }

    if ($workflowExists -eq $false) {
        _QCW-Log -Event 'QC_WORKFLOW_EXPECTED_MISSING' -Level 'Warning' -Message 'Expected current ProjectWise workflow was not validated.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_EXPECTED_MISSING' -Message 'Expected current ProjectWise workflow was not validated; continuing because workflow is informational.' -Data $data
    }
    if ($warnings.Count -gt 0) {
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message 'QC workflow informational validation completed with warnings.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_VALIDATION_WARNING' -Message 'QC workflow informational validation completed with warnings.' -Data $data
    }

    _QCW-Log -Event 'QC_WORKFLOW_VALIDATED' -Level 'Information' -Message 'QC workflow informational validation completed.' -Data $data
    return New-QCSuccessResult -Code 'QC_WORKFLOW_VALIDATED' -Message 'QC workflow informational validation completed.' -Data $data
}

function Test-QCWorkflowStateTransition {
    [CmdletBinding()]
    param(
        [hashtable]$Settings,
        [string]$CurrentStateName,
        [Parameter(Mandatory)]
        [string]$TargetStateName,
        [string]$WorkflowName,
        [switch]$ValidatePath
    )

    if (_QCW-IsNullOrWhiteSpace $WorkflowName -and $Settings) { $WorkflowName = [string]$Settings.expectedWorkflowName }
    $warnings = @()
    $links = @()
    $states = [System.Collections.Generic.List[string]]::new()
    $targetExists = $false
    $transitionValid = $null
    $workflowNameUsed = $WorkflowName

    $cmd = Get-Command -Name Get-PWWorkflowStateLinks -ErrorAction SilentlyContinue
    if (-not $cmd) {
        $warnings += 'Get-PWWorkflowStateLinks is unavailable; state transition validation cannot run.'
        $data = @{ workflowName = $WorkflowName; currentStateName = $CurrentStateName; targetStateName = $TargetStateName; targetStateExists = $null; transitionValid = $null; links = @(); warnings = @($warnings) }
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $warnings[0] -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_TRANSITION_UNKNOWN' -Message $warnings[0] -Data $data
    }

    $workflowCandidates = [System.Collections.Generic.List[string]]::new()
    foreach ($wn in @($WorkflowName, $(if ($Settings) { [string]$Settings.workflowName }), $(if ($Settings) { [string]$Settings.expectedWorkflowName }))) {
        if (-not (_QCW-IsNullOrWhiteSpace $wn) -and -not $workflowCandidates.Contains($wn.Trim())) { $workflowCandidates.Add($wn.Trim()) | Out-Null }
    }

    if ($workflowCandidates.Count -eq 0) {
        $warnings += 'No workflow name configured for Get-PWWorkflowStateLinks; using qcWorkflow.states fallback when strictMode is false.'
    }

    foreach ($wfTry in $workflowCandidates) {
        try {
            $tryLinks = @(_QCW-InvokeGetPWWorkflowStateLinks -WorkflowName $wfTry)
            if ($tryLinks.Count -gt 0) {
                $links = $tryLinks
                $workflowNameUsed = $wfTry
                break
            }
        } catch {
            $warnings += ('Get-PWWorkflowStateLinks failed for workflow ''' + $wfTry + ''': ' + $_.Exception.Message)
        }
    }

    $linkFormat = 'none'
    foreach ($link in @($links)) {
        foreach ($name in @('StateName','State','FromStateName','FromState','SourceStateName','SourceState','ToStateName','ToState','TargetStateName','TargetState','BeginState','EndState','BeginStateName','EndStateName')) {
            $v = _QCW-GetPropertyValue -Object $link -Names @($name)
            if (-not (_QCW-IsNullOrWhiteSpace $v)) {
                $sv = ([string]$v).Trim()
                $known = $false
                foreach ($existing in $states) { if (_QCW-StateNameEquals $existing $sv) { $known = $true; break } }
                if (-not $known) { $states.Add($sv) | Out-Null }
            }
        }
    }
    if (@($links).Count -gt 0 -and $states.Count -eq 0) {
        $first = $links[0]
        if ($null -ne (_QCW-GetPropertyValue -Object $first -Names @('Id')) -and
            ($null -ne (_QCW-GetPropertyValue -Object $first -Names @('iNext','iPrevious')))) {
            $linkFormat = 'id_chain'
            $warnings += 'Get-PWWorkflowStateLinks returned numeric state IDs (Id/iPrevious/iNext) without state names; use qcWorkflow.states or Set-PWDocumentState directly.'
        }
    } elseif ($states.Count -gt 0) {
        $linkFormat = 'named'
    }

    $targetExists = $false
    foreach ($s in $states) { if (_QCW-StateNameEquals $s $TargetStateName) { $targetExists = $true; break } }

    if ($targetExists -and (-not $ValidatePath -or (_QCW-IsNullOrWhiteSpace $CurrentStateName) -or (_QCW-StateNameEquals $CurrentStateName $TargetStateName))) {
        $transitionValid = $true
    } elseif ($targetExists -and $ValidatePath) {
        $transitionValid = $false
        foreach ($link in @($links)) {
            $from = _QCW-GetPropertyValue -Object $link -Names @('FromStateName','FromState','SourceStateName','SourceState','BeginState','BeginStateName','StateName','State')
            $to = _QCW-GetPropertyValue -Object $link -Names @('ToStateName','ToState','TargetStateName','TargetState','EndState','EndStateName')
            if (-not (_QCW-IsNullOrWhiteSpace $from) -and -not (_QCW-IsNullOrWhiteSpace $to) -and (_QCW-StateNameEquals $from $CurrentStateName) -and (_QCW-StateNameEquals $to $TargetStateName)) {
                $transitionValid = $true
                break
            }
        }
    }

    if ((-not $targetExists) -or ($linkFormat -eq 'id_chain')) {
        if (_QCW-TestTargetInConfiguredStates -Settings $Settings -TargetStateName $TargetStateName) {
            $targetExists = $true
            if ($ValidatePath -and -not (_QCW-StateNameEquals $CurrentStateName $TargetStateName)) {
                $transitionValid = $true
            }
            if ($linkFormat -eq 'id_chain') {
                $warnings += "Target state '$TargetStateName' validated against qcWorkflow.states (PW link rows use Id/iPrevious/iNext only)."
            } else {
                $warnings += "Target state '$TargetStateName' validated against qcWorkflow.states (workflow state links were empty or did not list this state)."
            }
        }
    }

    if (-not $targetExists) { $warnings += "Target workflow state '$TargetStateName' was not found in workflow state links." }
    elseif ($ValidatePath -and $transitionValid -eq $false) { $warnings += "No workflow state link from '$CurrentStateName' to '$TargetStateName' was found (Set-PWDocumentState may still succeed in ProjectWise)." }

    $data = @{ workflowName = $workflowNameUsed; currentStateName = $CurrentStateName; targetStateName = $TargetStateName; targetStateExists = $targetExists; transitionValid = $transitionValid; validatePath = [bool]$ValidatePath; linkFormat = $linkFormat; states = @($states); linkCount = @($links).Count; warnings = @($warnings) }
    if ($targetExists -and ($transitionValid -ne $false)) {
        _QCW-Log -Event 'QC_WORKFLOW_STATE_TRANSITION_VALID' -Level 'Information' -Message 'QC workflow target state/transition validated.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_TRANSITION_VALID' -Message 'QC workflow target state/transition validated.' -Data $data
    }

    _QCW-Log -Event 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Level 'Warning' -Message 'QC workflow target state/transition was not validated.' -Data $data
    return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Message 'QC workflow target state/transition was not validated.' -Data $data
}

function Set-PWQCWorkflowState {
    [CmdletBinding()]
    param(
        [hashtable]$Settings,
        [hashtable]$Context,
        [string]$StateName,
        [bool]$DryRun
    )

    $document = if ($Context -and $Context.ContainsKey('document')) { $Context.document } else { $null }
    $info = Get-PWDocumentWorkflowInfo -Document $document -Context $Context
    $data = @{ stateName = $StateName; currentStateName = $info.Data.stateName; transition = $null; planned = $false; changed = $false; warnings = @() }

    if (_QCW-IsNullOrWhiteSpace $StateName) {
        $data.warnings = @('Target workflow state is empty; no state change was made.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_MISSING' -Message $data.warnings[0] -Data $data
    }
    $transition = Test-QCWorkflowStateTransition -Settings $Settings -CurrentStateName $info.Data.stateName -TargetStateName $StateName -WorkflowName ([string]$Settings.expectedWorkflowName) -ValidatePath
    $data.transition = $transition
    if ($transition.Data -and $transition.Data.warnings) { $data.warnings = @($transition.Data.warnings) }
    if (-not $transition.IsSuccess -or ($transition.Data -and $transition.Data.transitionValid -eq $false)) {
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Message 'State change was not executed because transition validation failed.' -Data $data
    }
    if ($DryRun) {
        $data.planned = $true
        _QCW-Log -Event 'QC_WORKFLOW_STATE_PLANNED' -Level 'Information' -Message 'Dry-run: QC workflow state change planned.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_PLANNED' -Message 'Dry-run: would set document workflow state.' -Data $data
    }
    $configuredTargetOk = _QCW-TestTargetInConfiguredStates -Settings $Settings -TargetStateName $StateName
    if (-not ($transition.Data -and $transition.Data.targetStateExists -eq $true)) {
        if ($configuredTargetOk -and $document -and -not [bool]$Settings.strictMode) {
            $data.warnings = @("Proceeding with Set-PWDocumentState because target '$StateName' is listed in qcWorkflow.states (workflow link validation was inconclusive).")
            _QCW-Log -Event 'QC_WORKFLOW_STATE_CONFIGURED_FALLBACK' -Level 'Information' -Message $data.warnings[0] -Data $data
        } else {
            $data.warnings = @('State change was not executed because the target state could not be validated before a real write.')
            _QCW-Log -Event 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Level 'Warning' -Message $data.warnings[0] -Data $data
            return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Message $data.warnings[0] -Data $data
        }
    }

    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $document -or -not $cmd) {
        $data.warnings = @('ProjectWise state update requires a document object and Set-PWDocumentState; no state change was made.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_UNAVAILABLE' -Message $data.warnings[0] -Data $data
    }

    try {
        $args = @{}
        $docParam = _QCW-GetCommandParameterName -CommandName 'Set-PWDocumentState' -CandidateNames @('InputDocuments','InputDocument','Document')
        $stateParam = _QCW-GetCommandParameterName -CommandName 'Set-PWDocumentState' -CandidateNames @('StateName','State')
        if ($docParam) { $args[$docParam] = @($document) }
        if ($stateParam) { $args[$stateParam] = $StateName }
        if ((Get-Command -Name 'Set-PWDocumentState').Parameters.ContainsKey('ReturnBoolean')) { $args['ReturnBoolean'] = $true }
        if ($docParam -and $stateParam) { & $cmd @args -ErrorAction Stop | Out-Null }
        elseif ($stateParam) { & $cmd $document @args -ErrorAction Stop | Out-Null }
        else { & $cmd $document $StateName -ErrorAction Stop | Out-Null }
        $data.changed = $true
        _QCW-Log -Event 'QC_WORKFLOW_STATE_WRITE_SUCCESS' -Level 'Information' -Message 'QC workflow state write succeeded.' -Data $data
        $cfg = if ($Context -and $Context.ContainsKey('config')) { $Context.config } else { $null }

        $folderPath = ''
        $docName = ''
        $docGuid = ''
        if ($Context) {
            if ($Context.ContainsKey('documentPath') -and $Context.documentPath) {
                $dp = [string]$Context.documentPath
                if ($dp -match '\\') {
                    $folderPath = [System.IO.Path]::GetDirectoryName($dp)
                    if (-not $docName) { $docName = [System.IO.Path]::GetFileName($dp) }
                }
            }
            if ($Context.ContainsKey('job') -and $Context.job) {
                $job = $Context.job
                if (-not $folderPath -and $job.ContainsKey('sourceFolder')) { $folderPath = [string]$job.sourceFolder }
                if (-not $docName -and $job.ContainsKey('sourceName')) { $docName = [string]$job.sourceName }
            }
        }
        if ($document) {
            try {
                if (-not $docGuid -and $document.PSObject.Properties['DocumentGUID']) { $docGuid = [string]$document.DocumentGUID }
                if (-not $docName -and $document.PSObject.Properties['Name']) { $docName = [string]$document.Name }
                if (-not $folderPath -and $document.PSObject.Properties['FolderPath']) { $folderPath = [string]$document.FolderPath }
            } catch { }
        }

        if ($cfg -and -not (_QCW-IsNullOrWhiteSpace $folderPath) -and -not (_QCW-IsNullOrWhiteSpace $docName) `
            -and (Get-Command -Name 'Sync-PWAssociatedSheetMembersToWorkflowState' -ErrorAction SilentlyContinue)) {
            try {
                $sheetSync = Sync-PWAssociatedSheetMembersToWorkflowState -Config $cfg -FolderPath $folderPath `
                    -DocumentName $docName -DocumentGuid $docGuid -TargetStateName $StateName -DryRun:$DryRun `
                    -TriggerSource 'prepend_writeback'
                $data.sheetStateSync = $sheetSync
            } catch {
                $data.sheetStateSyncError = $_.Exception.Message
                _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message 'Associated sheet state sync failed after primary state write.' -Data @{
                    error = [string]$_.Exception.Message; folderPath = $folderPath; documentName = $docName; targetState = $StateName
                }
            }
        }

        if ($cfg) {
            $previousForTelemetry = [string]$info.Data.stateName
            if (_QCW-IsNullOrWhiteSpace $previousForTelemetry -and $Context) {
                if ($Context.ContainsKey('previousState') -and -not (_QCW-IsNullOrWhiteSpace $Context.previousState)) {
                    $previousForTelemetry = [string]$Context.previousState
                } elseif ($Context.ContainsKey('lifecycleState') -and -not (_QCW-IsNullOrWhiteSpace $Context.lifecycleState)) {
                    $previousForTelemetry = [string]$Context.lifecycleState
                }
            }
            $jobId = ''
            if ($Context -and $Context.ContainsKey('job') -and $Context.job -and $Context.job.ContainsKey('id')) {
                $jobId = [string]$Context.job.id
            }
            if (Get-Command -Name 'Invoke-QCSheetGroupWorkflowTransition' -ErrorAction SilentlyContinue) {
                if ($Context -and $data.sheetStateSync) { $Context['sheetStateSync'] = $data.sheetStateSync }
                Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $docGuid -TriggerDocumentName $docName `
                    -FolderPath $folderPath -SourceState $previousForTelemetry -TargetState $StateName `
                    -TransitionSource 'automation_prepend_completion' -JobId $jobId -JobType 'QC_PREPEND' `
                    -Context $Context -SuppressNotification -DryRun:$DryRun | Out-Null
            } elseif (Get-Command -Name 'Invoke-QCProcessorWorkflowStateTelemetry' -ErrorAction SilentlyContinue) {
                Invoke-QCProcessorWorkflowStateTelemetry -Config $cfg -Context $Context `
                    -PreviousState $previousForTelemetry -CurrentState $StateName -JobType 'QC_PREPEND' | Out-Null
            }
        }
        $notify = _QCW-InvokeStateChangeNotification -Config $cfg -Context $Context -PreviousState $info.Data.stateName -CurrentState $StateName -Document $document
        if ($notify) { $data.notification = $notify }
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_WRITE_SUCCESS' -Message 'QC workflow state write succeeded.' -Data $data
    } catch {
        $data.warnings = @($_.Exception.Message)
        _QCW-Log -Event 'QC_WORKFLOW_FAILURE' -Level 'Error' -Message 'QC workflow state write failed.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_FAILED' -Message 'QC workflow state write failed.' -Data $data
    }
}

function Set-PWQCAttributes {
    [CmdletBinding()]
    param(
        [hashtable]$Settings,
        [hashtable]$Context,
        [bool]$DryRun
    )

    $attributeMap = _QCW-ToHashtable $Settings.attributeMap
    $values = @{}
    if ($Context -and $Context.ContainsKey('attributes') -and $Context.attributes) {
        $normValues = _QCW-ToHashtable $Context.attributes
        if ($normValues) { $values = $normValues }
    }
    $filterResult = _QCW-FilterAttributeWritebackValues -Values $values -Settings $Settings
    $writebackValues = $filterResult.values
    $skippedKeys = @($filterResult.skippedKeys)

    $mapped = @{}
    if ($attributeMap) {
        foreach ($internalKey in $attributeMap.Keys) {
            if ($writebackValues.ContainsKey($internalKey) -and -not (_QCW-IsNullOrWhiteSpace $attributeMap[$internalKey])) {
                $mapped[[string]$attributeMap[$internalKey]] = $writebackValues[$internalKey]
            }
        }
    }

    $data = @{
        attributes = $mapped
        internalAttributes = $values
        writebackAttributes = $writebackValues
        skippedWritebackKeys = $skippedKeys
        planned = $false
        changed = $false
        warnings = @()
    }
    if ($skippedKeys.Count -gt 0) {
        _QCW-Log -Event 'QC_WORKFLOW_ATTRIBUTE_WRITEBACK_EXCLUDED' -Level 'Information' -Message 'Excluded attributes from ProjectWise writeback (DB/telemetry retains full context).' -Data @{
            skippedWritebackKeys = $skippedKeys
            writebackAttributes = $writebackValues
        }
    }
    if (-not $attributeMap -or $attributeMap.Keys.Count -eq 0) {
        $data.warnings = @('Attribute map is missing or empty; no attribute writeback was made.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_ATTRIBUTES_MAP_MISSING' -Message $data.warnings[0] -Data $data
    }
    if ($DryRun) {
        $data.planned = $true
        _QCW-Log -Event 'QC_WORKFLOW_ATTRIBUTES_PLANNED' -Level 'Information' -Message 'Dry-run: QC attribute writes planned.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ATTRIBUTES_PLANNED' -Message 'Dry-run: would write configured ProjectWise attributes.' -Data $data
    }

    if ($mapped.Keys.Count -eq 0) {
        _QCW-Log -Event 'QC_WORKFLOW_ATTRIBUTES_SKIPPED' -Level 'Information' -Message 'No ProjectWise attributes to write after writeback exclude filter.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ATTRIBUTES_SKIPPED' -Message 'No ProjectWise attributes to write after writeback exclude filter.' -Data $data
    }

    $document = if ($Context -and $Context.ContainsKey('document')) { $Context.document } else { $null }
    $cmd = Get-Command -Name 'Update-PWDocumentAttributes' -ErrorAction SilentlyContinue
    if (-not $document -or -not $cmd) {
        $data.warnings = @('ProjectWise attribute update requires a document object and Update-PWDocumentAttributes; no attributes were written.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_ATTRIBUTES_UNAVAILABLE' -Message $data.warnings[0] -Data $data
    }
    try {
        $args = @{}
        $docParam = _QCW-GetCommandParameterName -CommandName 'Update-PWDocumentAttributes' -CandidateNames @('InputDocuments','InputDocument','Document')
        $attrsParam = _QCW-GetCommandParameterName -CommandName 'Update-PWDocumentAttributes' -CandidateNames @('Attributes')
        if ($docParam) { $args[$docParam] = @($document) }
        if ($attrsParam) { $args[$attrsParam] = $mapped }
        if ((Get-Command -Name 'Update-PWDocumentAttributes').Parameters.ContainsKey('ReturnBoolean')) { $args['ReturnBoolean'] = $true }
        if ($docParam -and $attrsParam) { & $cmd @args -ErrorAction Stop | Out-Null }
        elseif ($attrsParam) { & $cmd $document @args -ErrorAction Stop | Out-Null }
        else { & $cmd $document $mapped -ErrorAction Stop | Out-Null }
        $data.changed = $true
        _QCW-Log -Event 'QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS' -Level 'Information' -Message 'QC attribute writeback succeeded.' -Data $data
        $cfg = if ($Context -and $Context.ContainsKey('config')) { $Context.config } else { $null }
        if ($cfg -and $mapped -and (Get-Command -Name 'Invoke-QCProcessorWorkflowAttributeTelemetry' -ErrorAction SilentlyContinue)) {
            Invoke-QCProcessorWorkflowAttributeTelemetry -Config $cfg -Context $Context -MappedAttributes $mapped | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS' -Message 'QC attribute writeback succeeded.' -Data $data
    } catch {
        $data.warnings = @($_.Exception.Message)
        _QCW-Log -Event 'QC_WORKFLOW_FAILURE' -Level 'Error' -Message 'QC attribute writeback failed.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_ATTRIBUTES_FAILED' -Message 'QC attribute writeback failed.' -Data $data
    }
}

function Advance-QCWorkflowCycleForRedlinesResubmit {
    <#
    Bumps the decimal sub-cycle (e.g. 1 -> 1.1 -> 1.2) when workflow reverses from
    Corrections Received back to Redlines Received within the same major QC cycle.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Settings,
        [Parameter(Mandatory)][hashtable]$Context,
        [hashtable]$Config = $null,
        [AllowEmptyString()][string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [bool]$DryRun = $false
    )

    $correctionsName = Get-QCWorkflowStateName -Settings $Settings -StateKey 'correctionsReceived'
    $redlinesName = Get-QCWorkflowStateName -Settings $Settings -StateKey 'redlinesReceived'
    if (_QCW-IsNullOrWhiteSpace $correctionsName) { $correctionsName = 'Corrections Received' }
    if (_QCW-IsNullOrWhiteSpace $redlinesName) { $redlinesName = 'Redlines Received' }

    if ([string]::Compare(([string]$PreviousState).Trim(), $correctionsName, $true) -ne 0) { return $Context }
    if ([string]::Compare(([string]$CurrentState).Trim(), $redlinesName, $true) -ne 0) { return $Context }

    $targets = _QCW-ResolveSheetCycleTargets -Context $Context
    $folderPath = [string]$targets.folderPath
    $sheetStem = [string]$targets.sheetStem
    $sourceDocGuid = [string]$targets.documentGuid

    $existing = $null
    if ($Config -and (Get-Command -Name 'Get-QCSheetIndexCycle' -ErrorAction SilentlyContinue)) {
        try {
            $existing = Get-QCSheetIndexCycle -Config $Config -DocumentGuid $sourceDocGuid -FolderPath $folderPath -SheetStem $sheetStem
        } catch { }
    }
    if (-not $existing -or ((_QCW-IsNullOrWhiteSpace $existing.cycleId) -and (_QCW-IsNullOrWhiteSpace $existing.cycleNumber))) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Warning' -Code 'QC_WORKFLOW_CYCLE_RESUBMIT_SKIPPED' `
                -Message 'Skipped redlines resubmit sub-cycle bump because no active QC cycle was found.' -Data @{
                folderPath = $folderPath
                sheetStem = $sheetStem
                documentGuid = $sourceDocGuid
                previousState = $PreviousState
                currentState = $CurrentState
            } | Out-Null
        }
        return $Context
    }

    $parsed = _QCW-ParseQCCycleNumber $existing.cycleNumber
    if ($parsed.major -le 0) {
        $parsed.major = 1
    }
    $nextMinor = $parsed.minor + 1
    $nextNumber = _QCW-FormatQCCycleNumber -Major $parsed.major -Minor $nextMinor
    $baseCycleId = _QCW-GetBaseQCCycleId ([string]$existing.cycleId)
    if (_QCW-IsNullOrWhiteSpace $baseCycleId) {
        $baseCycleId = [string]$existing.cycleId
    }
    $nextCycleId = _QCW-BuildQCCycleId -BaseCycleId $baseCycleId -CycleNumberDisplay $nextNumber

    $Context = _QCW-ApplyCycleToContext -Context $Context -CycleId $nextCycleId -CycleNumber $nextNumber

    if ($DryRun) { return $Context }

    if ($Config -and (Get-Command -Name 'Update-QCSheetIndexCycle' -ErrorAction SilentlyContinue)) {
        try {
            Update-QCSheetIndexCycle -Config $Config -CycleId $nextCycleId -CycleNumber $nextNumber `
                -DocumentGuid $sourceDocGuid -FolderPath $folderPath -SheetStem $sheetStem | Out-Null
        } catch { }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_WORKFLOW_CYCLE_RESUBMIT' `
            -Message 'Advanced QC sub-cycle for Corrections Received to Redlines Received transition.' -Data @{
            cycleId = $nextCycleId
            cycleNumber = $nextNumber
            previousCycleNumber = $existing.cycleNumber
            folderPath = $folderPath
            sheetStem = $sheetStem
            documentGuid = $sourceDocGuid
            previousState = $PreviousState
            currentState = $CurrentState
        } | Out-Null
    }

    return $Context
}

function Start-QCWorkflowCycleIfReadyForQc {
    <#
    Starts a new QC cycle when automation writeback lands on Ready for QC (prepend initialQcPdf).
    Persists cycleId/cycleNumber to context attributes, ProjectWise writeback, and sheet_index.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Settings,
        [Parameter(Mandatory)][hashtable]$Context,
        [hashtable]$Config = $null,
        [Parameter(Mandatory)][string]$TargetStateName,
        [bool]$DryRun = $false
    )

    $readyName = Get-QCWorkflowStateName -Settings $Settings -StateKey 'readyForQc'
    if (_QCW-IsNullOrWhiteSpace $readyName) { $readyName = 'Ready for QC' }
    if ([string]::Compare(([string]$TargetStateName).Trim(), $readyName, $true) -ne 0) {
        return $Context
    }

    $job = $null
    if ($Context -and $Context.ContainsKey('job') -and $Context.job) { $job = $Context.job }
    $prependTrigger = ''
    if ($Context -and $Context.ContainsKey('prependTrigger') -and $Context.prependTrigger) {
        $prependTrigger = [string]$Context.prependTrigger
    } elseif ($job -and $job.metadata -is [hashtable] -and $job.metadata.prependTrigger) {
        $prependTrigger = [string]$job.metadata.prependTrigger
    }

    $previousState = ''
    if ($Context -and $Context.ContainsKey('previousState') -and $Context.previousState) {
        $previousState = [string]$Context.previousState
    }
    $initiatedName = Get-QCWorkflowStateName -Settings $Settings -StateKey 'qcInitiated'
    if (_QCW-IsNullOrWhiteSpace $initiatedName) { $initiatedName = 'QC Initiated' }

    $shouldBump = $false
    if ($job -and $job.id) {
        if ($prependTrigger -eq 'initialQcPdf') { $shouldBump = $true }
        elseif (-not (_QCW-IsNullOrWhiteSpace $previousState) -and $previousState -ieq $initiatedName) { $shouldBump = $true }
    }
    if (-not $shouldBump) { return $Context }

    $folderPath = ''
    $sheetStem = ''
    $sourceDocGuid = ''
    if ($Context.ContainsKey('documentGuid') -and $Context.documentGuid) { $sourceDocGuid = [string]$Context.documentGuid }
    if ($job) {
        if (-not $folderPath -and $job.sourceFolder) { $folderPath = [string]$job.sourceFolder }
        if (-not $sourceDocGuid) {
            $md = _QCW-ToHashtable $job.metadata
            if ($md -and $md.triggerDocumentGuid) { $sourceDocGuid = [string]$md.triggerDocumentGuid }
            elseif ($md -and $md.documentGuid) { $sourceDocGuid = [string]$md.documentGuid }
        }
        if ($job.sourceName) {
            $srcName = [string]$job.sourceName
            if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                try { $sheetStem = [string](Get-PWSheetStemFromDocumentName -DocumentName $srcName) } catch { }
            }
            if (_QCW-IsNullOrWhiteSpace $sheetStem) {
                $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension($srcName)
            }
        }
    }
    $targets = _QCW-ResolveSheetCycleTargets -Context $Context
    if (_QCW-IsNullOrWhiteSpace $folderPath) { $folderPath = [string]$targets.folderPath }
    if (_QCW-IsNullOrWhiteSpace $sheetStem) { $sheetStem = [string]$targets.sheetStem }
    if (_QCW-IsNullOrWhiteSpace $sourceDocGuid) { $sourceDocGuid = [string]$targets.documentGuid }

    $parsedCurrent = _QCW-ParseQCCycleNumber $null
    if ($Config -and (Get-Command -Name 'Get-QCSheetIndexCycle' -ErrorAction SilentlyContinue)) {
        try {
            $existing = Get-QCSheetIndexCycle -Config $Config -DocumentGuid $sourceDocGuid -FolderPath $folderPath -SheetStem $sheetStem
            if ($existing -and -not (_QCW-IsNullOrWhiteSpace $existing.cycleNumber)) {
                $parsedCurrent = _QCW-ParseQCCycleNumber $existing.cycleNumber
            } elseif ($existing -and -not (_QCW-IsNullOrWhiteSpace $existing.cycleId)) {
                $parsedCurrent = _QCW-ParseQCCycleNumber 0
            }
        } catch { }
    }
    $nextMajor = [Math]::Max(0, $parsedCurrent.major) + 1
    $nextNumber = _QCW-FormatQCCycleNumber -Major $nextMajor -Minor 0
    $cycleId = [string]$job.id
    if (_QCW-IsNullOrWhiteSpace $cycleId) {
        $cycleId = 'qc_cycle_' + [guid]::NewGuid().ToString('n').Substring(0, 12)
    }
    $cycleId = _QCW-BuildQCCycleId -BaseCycleId $cycleId -CycleNumberDisplay $nextNumber

    $Context = _QCW-ApplyCycleToContext -Context $Context -CycleId $cycleId -CycleNumber $nextNumber

    if ($DryRun) { return $Context }

    if ($Config -and (Get-Command -Name 'Update-QCSheetIndexCycle' -ErrorAction SilentlyContinue)) {
        try {
            Update-QCSheetIndexCycle -Config $Config -CycleId $cycleId -CycleNumber $nextNumber `
                -DocumentGuid $sourceDocGuid -FolderPath $folderPath -SheetStem $sheetStem | Out-Null
        } catch { }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_WORKFLOW_CYCLE_STARTED' `
            -Message 'Started new QC cycle on automation Ready for QC writeback.' -Data @{
            cycleId = $cycleId
            cycleNumber = $nextNumber
            folderPath = $folderPath
            sheetStem = $sheetStem
            documentGuid = $sourceDocGuid
            prependTrigger = $prependTrigger
            previousState = $previousState
            targetState = $readyName
        } | Out-Null
    }

    return $Context
}

function Invoke-QCWorkflowWriteback {
    [CmdletBinding()]
    param(
        [hashtable]$Config,
        [hashtable]$Context
    )

    if ($Context -and $Config -and -not $Context.ContainsKey('config')) { $Context['config'] = $Config }

    $settings = Get-QCWorkflowSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        $data = @{ enabled = $false; actions = @(); warnings = @(); dryRun = $true }
        _QCW-Log -Event 'QC_WORKFLOW_DISABLED' -Level 'Information' -Message 'QC workflow writeback is disabled.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_DISABLED' -Message 'QC workflow writeback is disabled.' -Data $data
    }

    $validation = Test-QCWorkflowConfig -Config $Config
    if (-not $validation.IsSuccess) {
        _QCW-Log -Event 'QC_WORKFLOW_STRICT_FAILURE' -Level 'Error' -Message $validation.Message -Data @{ validation = $validation.Data }
        return $validation
    }

    $globalDryRun = $false
    if ($Config -and $Config.ContainsKey('dryRun')) { try { $globalDryRun = [bool]$Config.dryRun } catch { $globalDryRun = $false } }
    $dryRun = $globalDryRun -or [bool]$settings.dryRunWriteback
    if ($dryRun) {
        _QCW-Log -Event 'QC_WORKFLOW_DRYRUN' -Level 'Information' -Message 'QC workflow writeback is in dry-run mode.' -Data @{ globalDryRun = $globalDryRun; dryRunWriteback = [bool]$settings.dryRunWriteback }
    }

    $actions = [System.Collections.Generic.List[object]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    if ([bool]$settings.autoSetState -and ([string]$settings.mode) -ieq 'StateAndAttributes') {
        $assign = Ensure-PWQCWorkflowAssignment -Settings $settings -Context $Context -DryRun:$dryRun
        $actions.Add($assign) | Out-Null
        if ($assign.Data -and $assign.Data.warnings) { foreach ($w in @($assign.Data.warnings)) { if ($w) { $warnings.Add([string]$w) | Out-Null } } }
        if (-not $assign.IsSuccess -and [bool]$settings.strictMode) { return New-QCFailureResult -Code 'QC_WORKFLOW_STRICT_FAILURE' -Message $assign.Message -Data @{ actions = @($actions); warnings = @($warnings); settings = $settings; dryRun = $dryRun } }
    }

    if ([bool]$settings.autoSetState -and ([string]$settings.mode) -ieq 'StateAndAttributes') {
        $stateName = $null
        if ($Context -and $Context.ContainsKey('targetState') -and -not (_QCW-IsNullOrWhiteSpace $Context.targetState)) {
            $stateName = [string]$Context.targetState
        }
        if (_QCW-IsNullOrWhiteSpace $stateName) {
            if ($Context -and $Context.ContainsKey('resultStatus') -and ([string]$Context.resultStatus).ToLowerInvariant() -eq 'failed') {
                $stateName = [string]$settings.stateAfterFailedPrepend
            } else {
                $stateName = Resolve-QCWorkflowStateAfterPrepend -Settings $settings -Context $Context
            }
        }
        if (-not (_QCW-IsNullOrWhiteSpace $stateName)) {
            $Context = Start-QCWorkflowCycleIfReadyForQc -Settings $settings -Context $Context -Config $Config `
                -TargetStateName $stateName -DryRun:$dryRun
        }
        if (_QCW-IsNullOrWhiteSpace $stateName) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_EMPTY_STATE_GUARDED' `
                    -Message 'Skipped QC workflow writeback because resolved StateName was empty.' -Data @{
                    callSite = 'Invoke-PWQCWorkflowAutomation.stateName'
                    auditEventId = $null
                    documentName = if ($Context -and $Context.ContainsKey('documentName')) { [string]$Context.documentName } else { '' }
                    folderPath = if ($Context -and $Context.ContainsKey('sourceFolder')) { [string]$Context.sourceFolder } else { '' }
                    sourceVariableName = 'stateName'
                    sourceValue = $stateName
                    livePwState = if ($Context -and $Context.ContainsKey('previousState')) { [string]$Context.previousState } else { '' }
                    changedByUsername = ''
                } | Out-Null
            }
            $state = _QCW-NewWorkflowResult -IsSuccess $true -Code 'QC_WORKFLOW_STATE_EMPTY_SKIPPED' -Message 'Resolved workflow state was empty; no state write was attempted.' -Data @{ stateName = $stateName; changed = $false; planned = $false; warnings = @('Resolved workflow state was empty.') }
        } else {
            $state = Set-PWQCWorkflowState -Settings $settings -Context $Context -StateName $stateName -DryRun:$dryRun
        }
        $actions.Add($state) | Out-Null
        if ($state.Data -and $state.Data.warnings) { foreach ($w in @($state.Data.warnings)) { if ($w) { $warnings.Add([string]$w) | Out-Null } } }
        if (-not $state.IsSuccess -and [bool]$settings.strictMode) { return New-QCFailureResult -Code 'QC_WORKFLOW_STRICT_FAILURE' -Message $state.Message -Data @{ actions = @($actions); warnings = @($warnings); settings = $settings; dryRun = $dryRun } }
    }

    if ([bool]$settings.autoWriteAttributes) {
        $attrs = Set-PWQCAttributes -Settings $settings -Context $Context -DryRun:$dryRun
        $actions.Add($attrs) | Out-Null
        if ($attrs.Data -and $attrs.Data.warnings) { foreach ($w in @($attrs.Data.warnings)) { if ($w) { $warnings.Add([string]$w) | Out-Null } } }
        if (-not $attrs.IsSuccess -and [bool]$settings.strictMode) { return New-QCFailureResult -Code 'QC_WORKFLOW_STRICT_FAILURE' -Message $attrs.Message -Data @{ actions = @($actions); warnings = @($warnings); settings = $settings; dryRun = $dryRun } }
    }

    $summary = @{ enabled = $true; dryRun = $dryRun; actions = @($actions); warnings = @($warnings); mode = [string]$settings.mode; autoSetState = [bool]$settings.autoSetState; workflowName = [string]$settings.workflowName }
    foreach ($a in @($actions)) {
        if (-not $a) { continue }
        $ac = [string]$a.Code
        if ($ac -notmatch '^QC_WORKFLOW_STATE') { continue }
        $ad = _QCW-ToHashtable $a.Data
        if ($ad) {
            $summary['targetState'] = [string]$ad.stateName
            $summary['previousState'] = [string]$ad.currentStateName
            $summary['statePlanned'] = [bool]$ad.planned
            $summary['stateChanged'] = [bool]$ad.changed
            $summary['stateActionCode'] = $ac
        }
        break
    }
    _QCW-Log -Event 'QC_WORKFLOW_WRITEBACK_OK' -Level 'Information' -Message 'QC workflow writeback completed.' -Data $summary
    return New-QCSuccessResult -Code 'QC_WORKFLOW_WRITEBACK_OK' -Message 'QC workflow writeback completed.' -Data @{ enabled = $true; dryRun = $dryRun; actions = @($actions); warnings = @($warnings); settings = $settings }
}

Export-ModuleMember -Function Test-QCWorkflowConfig,Get-QCWorkflowSettings,Get-QCWorkflowAttributeWritebackExcludeDefaults,Get-QCWorkflowDeprecationWarnings,Get-QCWorkflowStateName,Normalize-QCPrependTriggerKey,Resolve-QCWorkflowStateAfterPrepend,Resolve-QCWorkflowAssignee,Get-PWDocumentWorkflowInfo,Ensure-PWQCWorkflowAssignment,Test-QCWorkflowStateTransition,Set-PWQCWorkflowState,Set-PWQCAttributes,Start-QCWorkflowCycleIfReadyForQc,Advance-QCWorkflowCycleForRedlinesResubmit,Invoke-QCWorkflowWriteback,Invoke-QCWorkflowStateChangeNotification
