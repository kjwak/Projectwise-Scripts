# QC.Workflow.psm1
# Responsibility: Configurable ProjectWise QC workflow/state/attribute writeback framework.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Logging.psm1') -Force -ErrorAction SilentlyContinue

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
        if (Get-Command -Name Write-QCLog -ErrorAction SilentlyContinue) {
            $payload = @{}
            if ($Data) { foreach ($k in $Data.Keys) { $payload[$k] = $Data[$k] } }
            $payload['event'] = $Event
            Write-QCLog -Level $Level -Message $Message -Data $payload | Out-Null
        }
    } catch { }
}

function _QCW-NewWorkflowResult([bool]$IsSuccess, [string]$Code, [string]$Message, [hashtable]$Data) {
    if ($IsSuccess) { return New-QCSuccessResult -Code $Code -Message $Message -Data $Data }
    return New-QCFailureResult -Code $Code -Message $Message -Data $Data
}

function _QCW-GetCommandParameterName([string]$CommandName, [string[]]$CandidateNames) {
    $cmd = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }
    foreach ($name in @($CandidateNames)) {
        if ($cmd.Parameters.ContainsKey($name)) { return $name }
    }
    return $null
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

function Get-QCWorkflowSettings {
    [CmdletBinding()]
    param([hashtable]$Config)

    $raw = @{}
    if ($Config -and $Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
        $norm = _QCW-ToHashtable $Config.qcWorkflow
        if ($norm) { $raw = $norm }
    }

    $defaults = @{
        enabled = $false
        strictMode = $false
        dryRunWriteback = $true
        workflowName = 'QC Review Workflow'
        expectedWorkflowName = 'QC Review Workflow'
        defaultStateAfterPrepend = 'QC Received'
        stateAfterSuccessfulPrepend = 'Redlines Issued'
        stateAfterFailedPrepend = 'Error / Needs Attention'
        autoAssignWorkflow = $true
        autoSetState = $true
        autoWriteAttributes = $true
        attributeMap = @{
            cycleId = 'QC_Cycle_ID'
            stage = 'QC_Stage'
            reviewer = 'QC_Reviewer'
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
        stageMap = @{
            red = @{ stageValue = 'Red'; stateName = 'Redlines Issued'; statusValue = 'Open' }
            green = @{ stageValue = 'Green'; stateName = 'Corrections Complete'; statusValue = 'Pending Backcheck' }
            blue = @{ stageValue = 'Blue'; stateName = 'Verified Closed'; statusValue = 'Closed' }
        }
    }

    $settings = @{}
    foreach ($k in $defaults.Keys) { $settings[$k] = $defaults[$k] }
    foreach ($k in $raw.Keys) {
        if ($k -eq 'attributeMap' -or $k -eq 'stageMap') { $settings[$k] = (_QCW-ToHashtable $raw[$k]) }
        else { $settings[$k] = $raw[$k] }
    }
    if (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName -and -not (_QCW-IsNullOrWhiteSpace $settings.workflowName)) {
        $settings.expectedWorkflowName = $settings.workflowName
    }
    if (_QCW-IsNullOrWhiteSpace $settings.workflowName -and -not (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName)) {
        $settings.workflowName = $settings.expectedWorkflowName
    }
    foreach ($boolKey in @('enabled','strictMode','dryRunWriteback','autoAssignWorkflow','autoSetState','autoWriteAttributes')) {
        try { $settings[$boolKey] = [bool]$settings[$boolKey] } catch { $settings[$boolKey] = [bool]$defaults[$boolKey] }
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
    $warnings = @()
    $critical = @()

    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'QC_WORKFLOW_CONFIG_DISABLED' -Message 'QC workflow writeback is disabled.' -Data @{ settings = $settings; warnings = @(); criticalIssues = @() }
    }

    if (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName) {
        $msg = 'qcWorkflow.enabled is true but qcWorkflow.expectedWorkflowName is empty.'
        $warnings += $msg
        $critical += $msg
    }
    $rawAttributeMap = $null
    if ($raw.ContainsKey('attributeMap')) { $rawAttributeMap = _QCW-ToHashtable $raw.attributeMap }
    if ([bool]$settings.autoWriteAttributes -and (-not $raw.ContainsKey('attributeMap') -or -not $rawAttributeMap -or $rawAttributeMap.Keys.Count -eq 0)) {
        $msg = 'qcWorkflow.autoWriteAttributes is true but qcWorkflow.attributeMap is missing or empty.'
        $warnings += $msg
        $critical += $msg
    }
    $stageMap = $null
    if ($raw.ContainsKey('stageMap')) { $stageMap = _QCW-ToHashtable $raw.stageMap }
    foreach ($stage in @('red','green','blue')) {
        if (-not $stageMap -or -not $stageMap.ContainsKey($stage)) {
            $warnings += "qcWorkflow.stageMap is missing '$stage'."
        }
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

    if (_QCW-IsNullOrWhiteSpace $expected) {
        $warnings += 'Expected workflow name is empty; inherited workflow validation cannot run.'
    } else {
        $wfCmd = Get-Command -Name Get-PWWorkflows -ErrorAction SilentlyContinue
        if ($wfCmd) {
            try {
                $workflows = @(Get-PWWorkflows -ErrorAction Stop)
                $workflowExists = [bool](@($workflows | Where-Object { _QCW-ObjectNameMatches -Object $_ -ExpectedName $expected -PropertyNames @('Name','WorkflowName','Workflow') }).Count -gt 0)
                if (-not $workflowExists) { $warnings += "Expected ProjectWise workflow '$expected' was not returned by Get-PWWorkflows." }
            } catch {
                $warnings += ('Get-PWWorkflows failed while validating expected workflow: ' + $_.Exception.Message)
            }
        } else {
            $warnings += 'Get-PWWorkflows is unavailable; expected workflow existence cannot be validated.'
        }
    }

    $matchesExpected = $false
    if (-not (_QCW-IsNullOrWhiteSpace $currentWorkflow) -and -not (_QCW-IsNullOrWhiteSpace $expected)) {
        $matchesExpected = ([string]$currentWorkflow).Trim() -eq $expected.Trim()
    }
    if (_QCW-IsNullOrWhiteSpace $currentWorkflow) {
        $warnings += 'Document object does not expose current workflow; folder-inherited workflow could not be validated from document properties.'
    } elseif (-not $matchesExpected) {
        $warnings += "Document current workflow '$currentWorkflow' does not match expected workflow '$expected'. Confirm the ProjectWise folder workflow inheritance configuration."
    }

    $data = @{
        expectedWorkflowName = $expected
        currentWorkflowName = $currentWorkflow
        workflowExists = $workflowExists
        documentMatchesExpectedWorkflow = $matchesExpected
        folderPath = $info.Data.folderPath
        dryRun = $DryRun
        changed = $false
        warnings = @($warnings)
    }

    if (($workflowExists -eq $false) -or (_QCW-IsNullOrWhiteSpace $expected)) {
        _QCW-Log -Event 'QC_WORKFLOW_EXPECTED_MISSING' -Level 'Warning' -Message 'Expected ProjectWise workflow was not validated.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_EXPECTED_MISSING' -Message 'Expected ProjectWise workflow was not validated.' -Data $data
    }
    if ($warnings.Count -gt 0) {
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message 'QC workflow inheritance validation completed with warnings.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_VALIDATION_WARNING' -Message 'QC workflow inheritance validation completed with warnings.' -Data $data
    }

    _QCW-Log -Event 'QC_WORKFLOW_VALIDATED' -Level 'Information' -Message 'QC workflow inheritance validated.' -Data $data
    return New-QCSuccessResult -Code 'QC_WORKFLOW_VALIDATED' -Message 'QC workflow inheritance validated.' -Data $data
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

    $cmd = Get-Command -Name Get-PWWorkflowStateLinks -ErrorAction SilentlyContinue
    if (-not $cmd) {
        $warnings += 'Get-PWWorkflowStateLinks is unavailable; state transition validation cannot run.'
        $data = @{ workflowName = $WorkflowName; currentStateName = $CurrentStateName; targetStateName = $TargetStateName; targetStateExists = $null; transitionValid = $null; links = @(); warnings = @($warnings) }
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $warnings[0] -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_TRANSITION_UNKNOWN' -Message $warnings[0] -Data $data
    }

    try {
        $args = @{}
        $wfParam = _QCW-GetCommandParameterName -CommandName 'Get-PWWorkflowStateLinks' -CandidateNames @('WorkflowName','Workflow')
        if ($wfParam -and -not (_QCW-IsNullOrWhiteSpace $WorkflowName)) { $args[$wfParam] = $WorkflowName }
        $links = @(& $cmd @args -ErrorAction Stop)
    } catch {
        $warnings += ('Get-PWWorkflowStateLinks failed: ' + $_.Exception.Message)
    }

    foreach ($link in @($links)) {
        foreach ($name in @('StateName','State','FromStateName','FromState','SourceStateName','SourceState','ToStateName','ToState','TargetStateName','TargetState')) {
            $v = _QCW-GetPropertyValue -Object $link -Names @($name)
            if (-not (_QCW-IsNullOrWhiteSpace $v) -and -not $states.Contains([string]$v)) { $states.Add([string]$v) | Out-Null }
        }
    }
    $targetExists = [bool](@($states | Where-Object { $_ -eq $TargetStateName }).Count -gt 0)

    if ($targetExists -and (-not $ValidatePath -or _QCW-IsNullOrWhiteSpace $CurrentStateName -or $CurrentStateName -eq $TargetStateName)) {
        $transitionValid = $true
    } elseif ($targetExists -and $ValidatePath) {
        $transitionValid = $false
        foreach ($link in @($links)) {
            $from = _QCW-GetPropertyValue -Object $link -Names @('FromStateName','FromState','SourceStateName','SourceState','StateName','State')
            $to = _QCW-GetPropertyValue -Object $link -Names @('ToStateName','ToState','TargetStateName','TargetState')
            if (-not (_QCW-IsNullOrWhiteSpace $from) -and -not (_QCW-IsNullOrWhiteSpace $to) -and ([string]$from -eq $CurrentStateName) -and ([string]$to -eq $TargetStateName)) {
                $transitionValid = $true
                break
            }
        }
    }

    if (-not $targetExists) { $warnings += "Target workflow state '$TargetStateName' was not found in workflow state links." }
    elseif ($ValidatePath -and $transitionValid -eq $false) { $warnings += "No workflow state link from '$CurrentStateName' to '$TargetStateName' was found." }

    $data = @{ workflowName = $WorkflowName; currentStateName = $CurrentStateName; targetStateName = $TargetStateName; targetStateExists = $targetExists; transitionValid = $transitionValid; validatePath = [bool]$ValidatePath; states = @($states); linkCount = @($links).Count; warnings = @($warnings) }
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
    if (-not ($transition.Data -and $transition.Data.targetStateExists -eq $true)) {
        $data.warnings = @('State change was not executed because the target state could not be validated before a real write.')
        _QCW-Log -Event 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_TRANSITION_INVALID' -Message $data.warnings[0] -Data $data
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
    $mapped = @{}
    if ($attributeMap) {
        foreach ($internalKey in $attributeMap.Keys) {
            if ($values.ContainsKey($internalKey) -and -not (_QCW-IsNullOrWhiteSpace $attributeMap[$internalKey])) {
                $mapped[[string]$attributeMap[$internalKey]] = $values[$internalKey]
            }
        }
    }

    $data = @{ attributes = $mapped; internalAttributes = $values; planned = $false; changed = $false; warnings = @() }
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
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS' -Message 'QC attribute writeback succeeded.' -Data $data
    } catch {
        $data.warnings = @($_.Exception.Message)
        _QCW-Log -Event 'QC_WORKFLOW_FAILURE' -Level 'Error' -Message 'QC attribute writeback failed.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_ATTRIBUTES_FAILED' -Message 'QC attribute writeback failed.' -Data $data
    }
}

function Invoke-QCWorkflowWriteback {
    [CmdletBinding()]
    param(
        [hashtable]$Config,
        [hashtable]$Context
    )

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

    if ([bool]$settings.autoAssignWorkflow) {
        $assign = Ensure-PWQCWorkflowAssignment -Settings $settings -Context $Context -DryRun:$dryRun
        $actions.Add($assign) | Out-Null
        if ($assign.Data -and $assign.Data.warnings) { foreach ($w in @($assign.Data.warnings)) { if ($w) { $warnings.Add([string]$w) | Out-Null } } }
        if (-not $assign.IsSuccess -and [bool]$settings.strictMode) { return New-QCFailureResult -Code 'QC_WORKFLOW_STRICT_FAILURE' -Message $assign.Message -Data @{ actions = @($actions); warnings = @($warnings); settings = $settings; dryRun = $dryRun } }
    }

    if ([bool]$settings.autoSetState) {
        $stageKey = $null
        try { if ($Context -and $Context.ContainsKey('stage') -and $Context.stage) { $stageKey = ([string]$Context.stage).Trim().ToLowerInvariant() } } catch { }
        $stateName = $null
        $stageMap = _QCW-ToHashtable $settings.stageMap
        if ($stageKey -and $stageMap -and $stageMap.ContainsKey($stageKey)) {
            $stageCfg = _QCW-ToHashtable $stageMap[$stageKey]
            if ($stageCfg -and $stageCfg.ContainsKey('stateName')) { $stateName = [string]$stageCfg.stateName }
        }
        if (_QCW-IsNullOrWhiteSpace $stateName) {
            if ($Context -and $Context.ContainsKey('resultStatus') -and ([string]$Context.resultStatus).ToLowerInvariant() -eq 'failed') { $stateName = [string]$settings.stateAfterFailedPrepend }
            else { $stateName = [string]$settings.stateAfterSuccessfulPrepend }
        }
        $state = Set-PWQCWorkflowState -Settings $settings -Context $Context -StateName $stateName -DryRun:$dryRun
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

    return New-QCSuccessResult -Code 'QC_WORKFLOW_WRITEBACK_OK' -Message 'QC workflow writeback completed.' -Data @{ enabled = $true; dryRun = $dryRun; actions = @($actions); warnings = @($warnings); settings = $settings }
}

Export-ModuleMember -Function Test-QCWorkflowConfig,Get-QCWorkflowSettings,Get-PWDocumentWorkflowInfo,Ensure-PWQCWorkflowAssignment,Test-QCWorkflowStateTransition,Set-PWQCWorkflowState,Set-PWQCAttributes,Invoke-QCWorkflowWriteback
