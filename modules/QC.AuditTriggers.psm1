# QC.AuditTriggers.psm1
# Responsibility: Audit-driven QC workflow state/attribute triggers (telemetry + notifications).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
if (-not (Get-Command -Name 'Write-QCDocumentStateHistoryRow' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'Core.Database.psm1') -Force
}
if (-not (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'QC.Notifications.psm1') -Force -ErrorAction SilentlyContinue
}

function _QCAT-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCAT-NormalizeValue([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return ([string]$Value).Trim()
}

function Test-QCIsQcPdfDocumentName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DocumentName)
    return ([string]$DocumentName -match '(?i)-qc\.pdf$')
}

function Get-QCAuditWorkflowTriggerSettings {
    <#
    .SYNOPSIS
    Resolves auditPoller.workflowTriggers settings (defaults enable telemetry; notifications follow notifications.enabled).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $defaults = @{
        enabled                = $true
        recordStateHistory     = $true
        recordAttributeHistory = $true
        recordTransitions      = $true
        recordFromProcessor    = $true
        notifyOnStateChange    = $true
        qcPdfNotificationsOnly = $true
    }
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = _QCAT-ToHashtable $Config.auditPoller
            if ($ap -and $ap.ContainsKey('workflowTriggers') -and $ap.workflowTriggers) {
                $wt = _QCAT-ToHashtable $ap.workflowTriggers
                if ($wt) {
                    foreach ($k in $wt.Keys) {
                        if ($defaults.ContainsKey($k)) {
                            try { $defaults[$k] = [bool]$wt[$k] } catch { }
                        }
                    }
                }
            }
        }
    } catch { }
    return $defaults
}

function _QCAT-BuildNotificationDocument {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$DocumentName,
        [string]$DocumentGuid,
        [hashtable]$Attributes = @{}
    )
    $notifCfg = $null
    try {
        if ($Config.ContainsKey('notifications') -and $Config.notifications) {
            $notifCfg = _QCAT-ToHashtable $Config.notifications
        }
    } catch { }
    $attrCfg = @{}
    if ($notifCfg -and $notifCfg.ContainsKey('attributes')) {
        $attrCfg = _QCAT-ToHashtable $notifCfg.attributes
        if (-not $attrCfg) { $attrCfg = @{} }
    }
    $reviewerField = if ($attrCfg.reviewerEmailField) { [string]$attrCfg.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $designerField = if ($attrCfg.designerEmailField) { [string]$attrCfg.designerEmailField } else { 'EM_Designer_Email' }

    $doc = [ordered]@{
        Name         = $DocumentName
        DocumentName = $DocumentName
        DocumentGUID = $DocumentGuid
        FolderPath   = $FolderPath
    }
    if ($Attributes -and $Attributes.Count -gt 0) {
        foreach ($k in $Attributes.Keys) { $doc[$k] = $Attributes[$k] }
        if ($Attributes.ContainsKey($reviewerField)) { $doc[$reviewerField] = $Attributes[$reviewerField] }
        if ($Attributes.ContainsKey($designerField)) { $doc[$designerField] = $Attributes[$designerField] }
    }
    return [pscustomobject]$doc
}

function Invoke-QCAuditWorkflowStateChangeTriggers {
    <#
    .SYNOPSIS
    Records workflow state transitions from audit events and optionally sends QC PDF notifications.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [object]$Document = $null,
        [hashtable]$PwAttributes = $null,
        [bool]$DryRun = $false,
        [string]$AuditActionName = 'DOCUMENT_STATE',
        [Nullable[int]]$ChangedByUser = $null
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled) { return }

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    if ($prev -eq $curr) { return }

    $shouldNotify = [bool]$settings.notifyOnStateChange
    if ($shouldNotify -and [bool]$settings.qcPdfNotificationsOnly -and -not (Test-QCIsQcPdfDocumentName -DocumentName $DocumentName)) {
        $shouldNotify = $false
    }

    if ([bool]$settings.recordStateHistory) {
        Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -EventType 'STATE_CHANGE' -OldValue $prev -NewValue $curr `
            -FieldName 'pw_state_name' -ChangedByUser $ChangedByUser | Out-Null
    }

    $transitionId = $null
    if ([bool]$settings.recordTransitions) {
        $tr = Write-QCTransitionEvent -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -TransitionType 'STATE_CHANGE' -FromValue $prev -ToValue $curr `
            -JobType 'audit_trigger'
        if ($tr.IsSuccess -and $tr.Data -and $null -ne $tr.Data.transitionId) {
            try { $transitionId = [int]$tr.Data.transitionId } catch { $transitionId = $null }
        }
    }

    if (-not $shouldNotify) { return }
    if (-not (Get-Command -Name 'Invoke-QCNotificationForStateChange' -ErrorAction SilentlyContinue)) { return }

    if (-not $Document) {
        $attrs = @{}
        if ($PwAttributes) { $attrs = $PwAttributes }
        $Document = _QCAT-BuildNotificationDocument -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -Attributes $attrs
    }

    if ($DryRun) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_NOTIFY_PLANNED' -Message 'QC notification planned (dry-run).' -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; previousState = $prev; currentState = $curr
            }
        }
        return
    }

    try {
        $notif = Invoke-QCNotificationForStateChange -Config $Config -PreviousState $prev -CurrentState $curr `
            -Document $Document -DocumentName $DocumentName -DocumentGuid $DocumentGuid `
            -DocumentPath ($FolderPath + '\' + $DocumentName)
        if ($transitionId -and $notif -and $notif.IsSuccess -and $notif.Data) {
            $nid = $null
            if ($notif.Data.ContainsKey('dedupeKey')) { $nid = [string]$notif.Data.dedupeKey }
            elseif ($notif.Data.dedupeKey) { $nid = [string]$notif.Data.dedupeKey }
            $sent = $false
            try { if ($notif.Data.success -eq $true) { $sent = $true } } catch { }
            if ($sent) {
                Update-QCTransitionEventNotification -Config $Config -TransitionId $transitionId -NotificationSent $true -NotificationId $nid
            }
        }
    } catch {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_NOTIFY_FAILED' -Message $_.Exception.Message -Data @{
                documentGuid = $DocumentGuid; documentName = $DocumentName; previousState = $prev; currentState = $curr
            }
        }
    }
}

function _QCAT-GetContextDocumentIdentity {
    param([hashtable]$Context)
    $document = $null
    if ($Context -and $Context.ContainsKey('document')) { $document = $Context.document }
    $docGuid = ''
    $docName = ''
    $folderPath = ''
    if ($document) {
        try {
            if ($document.PSObject.Properties['DocumentGUID'] -and $document.DocumentGUID) { $docGuid = [string]$document.DocumentGUID }
            elseif ($document.PSObject.Properties['DocumentGuid'] -and $document.DocumentGuid) { $docGuid = [string]$document.DocumentGuid }
        } catch { }
        try {
            if ($document.PSObject.Properties['Name'] -and $document.Name) { $docName = [string]$document.Name }
            elseif ($document.PSObject.Properties['DocumentName'] -and $document.DocumentName) { $docName = [string]$document.DocumentName }
        } catch { }
        try {
            if ($document.PSObject.Properties['FolderPath'] -and $document.FolderPath) { $folderPath = [string]$document.FolderPath }
        } catch { }
    }
    if ($Context) {
        if (-not $docGuid -and $Context.ContainsKey('documentGuid')) { $docGuid = [string]$Context.documentGuid }
        if (-not $docName -and $Context.ContainsKey('documentName')) { $docName = [string]$Context.documentName }
        if (-not $folderPath) {
            foreach ($k in @('folderPath', 'sourceFolder', 'incomingFolderPath')) {
                if ($Context.ContainsKey($k) -and $Context[$k]) { $folderPath = [string]$Context[$k]; break }
            }
        }
    }
    $job = $null
    if ($Context -and $Context.ContainsKey('job')) { $job = _QCAT-ToHashtable $Context.job }
    if ($job) {
        if (-not $docGuid -and $job.ContainsKey('documentGuid')) { $docGuid = [string]$job.documentGuid }
        if (-not $docName -and $job.ContainsKey('fileName')) { $docName = [string]$job.fileName }
        if (-not $folderPath -and $job.ContainsKey('sourceFolder')) { $folderPath = [string]$job.sourceFolder }
        if (-not $folderPath -and $job.ContainsKey('metadata')) {
            $meta = _QCAT-ToHashtable $job.metadata
            if ($meta) {
                if (-not $docGuid -and $meta.ContainsKey('documentId')) { $docGuid = [string]$meta.documentId }
                if (-not $folderPath -and $meta.ContainsKey('folderPath')) { $folderPath = [string]$meta.folderPath }
            }
        }
    }
    return @{ documentGuid = $docGuid; documentName = $docName; folderPath = $folderPath; job = $job }
}

function Invoke-QCProcessorWorkflowStateTelemetry {
    <#
    .SYNOPSIS
    Records processor-driven workflow state changes (QC_PREPEND / comment sync) without sending email.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Context,
        [string]$PreviousState = '',
        [Parameter(Mandatory)][string]$CurrentState,
        [string]$JobType = 'processor'
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled -or -not [bool]$settings.recordFromProcessor) { return }

    $prev = _QCAT-NormalizeValue $PreviousState
    $curr = _QCAT-NormalizeValue $CurrentState
    if ($prev -eq $curr) { return }

    $id = _QCAT-GetContextDocumentIdentity -Context $Context
    if ([string]::IsNullOrWhiteSpace($id.documentGuid)) { return }

    $jobId = ''
    if ($id.job -and $id.job.ContainsKey('id')) { $jobId = [string]$id.job.id }

    if ([bool]$settings.recordStateHistory) {
        Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
            -FolderPath $id.folderPath -EventType 'STATE_CHANGE' -OldValue $prev -NewValue $curr `
            -FieldName 'pw_state_name' | Out-Null
    }
    if ([bool]$settings.recordTransitions) {
        Write-QCTransitionEvent -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
            -FolderPath $id.folderPath -TransitionType 'STATE_CHANGE' -FromValue $prev -ToValue $curr `
            -JobId $jobId -JobType $JobType | Out-Null
    }
}

function Invoke-QCProcessorWorkflowAttributeTelemetry {
    <#
    .SYNOPSIS
    Records attribute writeback from QC_PREPEND processor (new values only; audit path has before/after).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][hashtable]$MappedAttributes
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled -or -not [bool]$settings.recordFromProcessor) { return }
    if (-not [bool]$settings.recordAttributeHistory) { return }
    if (-not $MappedAttributes -or $MappedAttributes.Keys.Count -eq 0) { return }

    $id = _QCAT-GetContextDocumentIdentity -Context $Context
    if ([string]::IsNullOrWhiteSpace($id.documentGuid)) { return }

    $jobId = ''
    if ($id.job -and $id.job.ContainsKey('id')) { $jobId = [string]$id.job.id }

    foreach ($field in @($MappedAttributes.Keys)) {
        $newVal = _QCAT-NormalizeValue ([string]$MappedAttributes[$field])
        if ([string]::IsNullOrWhiteSpace($newVal)) { continue }
        Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
            -FolderPath $id.folderPath -EventType 'ATTR_CHANGE' -OldValue '' -NewValue $newVal `
            -FieldName ([string]$field) | Out-Null
        if ([bool]$settings.recordTransitions) {
            Write-QCTransitionEvent -Config $Config -DocumentGuid $id.documentGuid -DocumentName $id.documentName `
                -FolderPath $id.folderPath -TransitionType 'ATTR_CHANGE' -FromValue '' -ToValue $newVal `
                -JobId $jobId -JobType 'QC_PREPEND' | Out-Null
        }
    }
}

function Invoke-QCAuditWorkflowAttributeChangeTriggers {
    <#
    .SYNOPSIS
    Records attribute field transitions detected after DOCUMENT_ATTR audit events (sheet_index diff).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][hashtable]$FieldChanges,
        [Nullable[int]]$ChangedByUser = $null
    )

    $settings = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$settings.enabled) { return }
    if (-not $FieldChanges -or $FieldChanges.Keys.Count -eq 0) { return }

    foreach ($field in @($FieldChanges.Keys)) {
        $change = _QCAT-ToHashtable $FieldChanges[$field]
        if (-not $change) { continue }
        $oldVal = _QCAT-NormalizeValue ([string]$change.oldValue)
        $newVal = _QCAT-NormalizeValue ([string]$change.newValue)
        if ($oldVal -eq $newVal) { continue }

        if ([bool]$settings.recordAttributeHistory) {
            Write-QCDocumentStateHistoryRow -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
                -FolderPath $FolderPath -EventType 'ATTR_CHANGE' -OldValue $oldVal -NewValue $newVal `
                -FieldName ([string]$field) -ChangedByUser $ChangedByUser | Out-Null
        }

        if ([bool]$settings.recordTransitions) {
            Write-QCTransitionEvent -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
                -FolderPath $FolderPath -TransitionType 'ATTR_CHANGE' -FromValue $oldVal -ToValue $newVal `
                -JobType 'audit_trigger' | Out-Null
        }
    }
}

Export-ModuleMember -Function Get-QCAuditWorkflowTriggerSettings, Test-QCIsQcPdfDocumentName, `
    Invoke-QCAuditWorkflowStateChangeTriggers, Invoke-QCAuditWorkflowAttributeChangeTriggers, `
    Invoke-QCProcessorWorkflowStateTelemetry, Invoke-QCProcessorWorkflowAttributeTelemetry
