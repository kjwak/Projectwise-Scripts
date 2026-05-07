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

function _QCW-GetValue([hashtable]$Hash, [string]$Key, [object]$DefaultValue) {
    if ($Hash -and $Hash.ContainsKey($Key) -and $null -ne $Hash[$Key]) { return $Hash[$Key] }
    return $DefaultValue
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

    if (_QCW-IsNullOrWhiteSpace $settings.workflowName) {
        $msg = 'qcWorkflow.enabled is true but qcWorkflow.workflowName is empty.'
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

    $workflowName = $null
    $stateName = $null
    foreach ($name in @('WorkflowName','Workflow','WorkflowStateName')) {
        try { if ($Document -and $Document.PSObject.Properties[$name] -and $Document.$name) { $workflowName = [string]$Document.$name; break } } catch { }
    }
    foreach ($name in @('StateName','DocumentState','WorkflowState','CurrentState')) {
        try { if ($Document -and $Document.PSObject.Properties[$name] -and $Document.$name) { $stateName = [string]$Document.$name; break } } catch { }
    }

    return New-QCSuccessResult -Code 'QC_WORKFLOW_INFO' -Message 'ProjectWise workflow info resolved from available document properties.' -Data @{
        workflowName = $workflowName
        stateName = $stateName
        document = $Document
        documentPath = if ($Context -and $Context.ContainsKey('documentPath')) { $Context.documentPath } else { $null }
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
    $info = Get-PWDocumentWorkflowInfo -Document $document -Context $Context
    $alreadyAssigned = (-not (_QCW-IsNullOrWhiteSpace $info.Data.workflowName)) -and (([string]$info.Data.workflowName) -eq ([string]$Settings.workflowName))
    $data = @{ workflowName = [string]$Settings.workflowName; alreadyAssigned = $alreadyAssigned; planned = $false; changed = $false; warnings = @() }

    if ($alreadyAssigned) { return New-QCSuccessResult -Code 'QC_WORKFLOW_ASSIGN_ALREADY_SET' -Message 'Document is already assigned to the configured QC workflow.' -Data $data }
    if ($DryRun) {
        $data.planned = $true
        _QCW-Log -Event 'QC_WORKFLOW_ASSIGN_PLANNED' -Level 'Information' -Message 'Dry-run: QC workflow assignment planned.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ASSIGN_PLANNED' -Message 'Dry-run: would assign document to configured QC workflow.' -Data $data
    }

    # TODO(ProjectWise): Verify the exact site-approved workflow assignment cmdlet/signature.
    # This isolated wrapper intentionally avoids hard-coding workflow calls in QC_PREPEND.
    $cmd = Get-Command -Name 'Set-PWDocumentWorkflow' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        $data.warnings = @('ProjectWise workflow assignment cmdlet Set-PWDocumentWorkflow is not available; no workflow assignment was made.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ASSIGN_UNAVAILABLE' -Message $data.warnings[0] -Data $data
    }
    try {
        & $cmd -Document $document -WorkflowName ([string]$Settings.workflowName) -ErrorAction Stop | Out-Null
        $data.changed = $true
        _QCW-Log -Event 'QC_WORKFLOW_ASSIGN_SUCCESS' -Level 'Information' -Message 'QC workflow assignment succeeded.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ASSIGN_SUCCESS' -Message 'QC workflow assignment succeeded.' -Data $data
    } catch {
        $data.warnings = @($_.Exception.Message)
        _QCW-Log -Event 'QC_WORKFLOW_FAILURE' -Level 'Error' -Message 'QC workflow assignment failed.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_ASSIGN_FAILED' -Message 'QC workflow assignment failed.' -Data $data
    }
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
    $data = @{ stateName = $StateName; planned = $false; changed = $false; warnings = @() }
    if (_QCW-IsNullOrWhiteSpace $StateName) {
        $data.warnings = @('Target workflow state is empty; no state change was made.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_MISSING' -Message $data.warnings[0] -Data $data
    }
    if ($DryRun) {
        $data.planned = $true
        _QCW-Log -Event 'QC_WORKFLOW_STATE_PLANNED' -Level 'Information' -Message 'Dry-run: QC workflow state change planned.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_PLANNED' -Message 'Dry-run: would set document workflow state.' -Data $data
    }

    # TODO(ProjectWise): Verify the exact site-approved state transition cmdlet/signature.
    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $cmd) { $cmd = Get-Command -Name 'Set-PWDocumentWorkflowState' -ErrorAction SilentlyContinue }
    if (-not $cmd) {
        $data.warnings = @('ProjectWise state cmdlet Set-PWDocumentState/Set-PWDocumentWorkflowState is not available; no state change was made.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_UNAVAILABLE' -Message $data.warnings[0] -Data $data
    }
    try {
        & $cmd -Document $document -StateName $StateName -ErrorAction Stop | Out-Null
        $data.changed = $true
        _QCW-Log -Event 'QC_WORKFLOW_STATE_SUCCESS' -Level 'Information' -Message 'QC workflow state change succeeded.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_STATE_SUCCESS' -Message 'QC workflow state change succeeded.' -Data $data
    } catch {
        $data.warnings = @($_.Exception.Message)
        _QCW-Log -Event 'QC_WORKFLOW_FAILURE' -Level 'Error' -Message 'QC workflow state change failed.' -Data $data
        return _QCW-NewWorkflowResult -IsSuccess (-not ([bool]$Settings.strictMode)) -Code 'QC_WORKFLOW_STATE_FAILED' -Message 'QC workflow state change failed.' -Data $data
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

    # TODO(ProjectWise): Verify environment-specific attribute storage. This default path sets matching
    # document properties and calls Update-PWDocumentProperties when available.
    $document = if ($Context -and $Context.ContainsKey('document')) { $Context.document } else { $null }
    $cmd = Get-Command -Name 'Update-PWDocumentProperties' -ErrorAction SilentlyContinue
    if (-not $document -or -not $cmd) {
        $data.warnings = @('ProjectWise attribute update requires a document object and Update-PWDocumentProperties; no attributes were written.')
        _QCW-Log -Event 'QC_WORKFLOW_WARNING' -Level 'Warning' -Message $data.warnings[0] -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ATTRIBUTES_UNAVAILABLE' -Message $data.warnings[0] -Data $data
    }
    try {
        foreach ($name in $mapped.Keys) {
            if ($document.PSObject.Properties[$name]) { $document.$name = $mapped[$name] }
            else { Add-Member -InputObject $document -NotePropertyName $name -NotePropertyValue $mapped[$name] -Force }
        }
        & $cmd $document -ErrorAction Stop | Out-Null
        $data.changed = $true
        _QCW-Log -Event 'QC_WORKFLOW_ATTRIBUTES_SUCCESS' -Level 'Information' -Message 'QC attribute writeback succeeded.' -Data $data
        return New-QCSuccessResult -Code 'QC_WORKFLOW_ATTRIBUTES_SUCCESS' -Message 'QC attribute writeback succeeded.' -Data $data
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

Export-ModuleMember -Function Test-QCWorkflowConfig,Get-QCWorkflowSettings,Get-PWDocumentWorkflowInfo,Ensure-PWQCWorkflowAssignment,Set-PWQCWorkflowState,Set-PWQCAttributes,Invoke-QCWorkflowWriteback
