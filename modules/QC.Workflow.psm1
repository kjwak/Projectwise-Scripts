# QC.Workflow.psm1
# Responsibility: Configurable ProjectWise QC workflow/state/attribute writeback framework.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'Core.Logging.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.Notifications.psm1') -Force -ErrorAction SilentlyContinue
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

function _QCW-InvokeStateChangeNotification {
    param(
        [hashtable]$Config,
        [hashtable]$Context,
        [string]$PreviousState,
        [string]$CurrentState,
        [object]$Document
    )

    if (-not $Config) { return $null }
    if (-not (Get-Command -Name Invoke-QCNotificationForStateChange -ErrorAction SilentlyContinue)) { return $null }
    $job = $null
    if ($Context -and $Context.ContainsKey('job')) { $job = $Context.job }
    try {
        return Invoke-QCNotificationForStateChange -Config $Config -PreviousState $PreviousState -CurrentState $CurrentState `
            -Document $Document -Job $job
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
        readyForQc = 'Ready for QC'
        reviewInProgress = 'Review In Progress'
        redlinesIssued = 'Redlines Issued'
        correctionsInProgress = 'Corrections In Progress'
        verificationInProgress = 'Verification In Progress'
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

    return @($warnings)
}

function _QCW-DefaultPrependStateTriggers {
    $states = _QCW-DefaultWorkflowStates
    return @{
        initialQcPdf = [string]$states.readyForQc
        reviewerRedlineUpdate = [string]$states.redlinesIssued
        designerCorrectionComplete = [string]$states.verificationInProgress
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
        'reviewerredlineupdate' { return 'reviewerRedlineUpdate' }
        'reviewerredline' { return 'reviewerRedlineUpdate' }
        'redlineupdate' { return 'reviewerRedlineUpdate' }
        'redlinesissued' { return 'reviewerRedlineUpdate' }
        'designercorrectioncomplete' { return 'designerCorrectionComplete' }
        'correctioncomplete' { return 'designerCorrectionComplete' }
        'correctionscomplete' { return 'designerCorrectionComplete' }
        default {
            if ($v -eq 'initialQcPdf' -or $v -eq 'reviewerRedlineUpdate' -or $v -eq 'designerCorrectionComplete') { return $v }
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
    $defaults = _QCW-DefaultWorkflowStates
    if ($defaults.ContainsKey($StateKey)) { return [string]$defaults[$StateKey] }
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
    $ready = [string]$states.readyForQc
    $review = [string]$states.reviewInProgress
    $redlinesIssued = [string]$states.redlinesIssued
    $corrections = [string]$states.correctionsInProgress
    $verification = [string]$states.verificationInProgress
    $complete = [string]$states.complete

    if ($state -eq $complete) { return $null }

    if ($state -eq $production -or $state -eq $redlinesIssued -or $state -eq $corrections) {
        if (-not (_QCW-IsNullOrWhiteSpace $DesignerEmail)) { return [string]$DesignerEmail }
        return $null
    }
    if ($state -in @($ready, $review, $verification)) {
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
        if ($raw.ContainsKey('receivedStateName') -and $raw.receivedStateName) { $settings.states.readyForQc = [string]$raw.receivedStateName }
        if ($raw.ContainsKey('correctionsInProgressStateName') -and $raw.correctionsInProgressStateName) { $settings.states.correctionsInProgress = [string]$raw.correctionsInProgressStateName }
        if ($raw.ContainsKey('backcheckInProgressStateName') -and $raw.backcheckInProgressStateName) { $settings.states.verificationInProgress = [string]$raw.backcheckInProgressStateName }
        if ($raw.ContainsKey('errorStateName') -and $raw.errorStateName) { $settings.states.error = [string]$raw.errorStateName }
    }

    if (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName -and -not (_QCW-IsNullOrWhiteSpace $settings.workflowName)) {
        $settings.expectedWorkflowName = $settings.workflowName
    }
    if (_QCW-IsNullOrWhiteSpace $settings.workflowName -and -not (_QCW-IsNullOrWhiteSpace $settings.expectedWorkflowName)) {
        $settings.workflowName = $settings.expectedWorkflowName
    }
    foreach ($boolKey in @('enabled','strictMode','dryRunWriteback','autoSetState','autoWriteAttributes')) {
        try { $settings[$boolKey] = [bool]$settings[$boolKey] } catch { }
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

    foreach ($wfTry in $workflowCandidates) {
        try {
            $args = @{}
            $wfParam = _QCW-GetCommandParameterName -CommandName 'Get-PWWorkflowStateLinks' -CandidateNames @('WorkflowName','Workflow')
            if ($wfParam) { $args[$wfParam] = $wfTry }
            $tryLinks = @(& $cmd @args -ErrorAction Stop)
            if ($tryLinks.Count -gt 0) {
                $links = $tryLinks
                $workflowNameUsed = $wfTry
                break
            }
        } catch {
            $warnings += ('Get-PWWorkflowStateLinks failed for workflow ''' + $wfTry + ''': ' + $_.Exception.Message)
        }
    }

    foreach ($link in @($links)) {
        foreach ($name in @('StateName','State','FromStateName','FromState','SourceStateName','SourceState','ToStateName','ToState','TargetStateName','TargetState')) {
            $v = _QCW-GetPropertyValue -Object $link -Names @($name)
            if (-not (_QCW-IsNullOrWhiteSpace $v) -and -not $states.Contains([string]$v)) { $states.Add([string]$v) | Out-Null }
        }
    }
    $targetExists = [bool](@($states | Where-Object { $_ -eq $TargetStateName }).Count -gt 0)

    if ($targetExists -and (-not $ValidatePath -or (_QCW-IsNullOrWhiteSpace $CurrentStateName) -or $CurrentStateName -eq $TargetStateName)) {
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

    $data = @{ workflowName = $workflowNameUsed; currentStateName = $CurrentStateName; targetStateName = $TargetStateName; targetStateExists = $targetExists; transitionValid = $transitionValid; validatePath = [bool]$ValidatePath; states = @($states); linkCount = @($links).Count; warnings = @($warnings) }
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
        $cfg = if ($Context -and $Context.ContainsKey('config')) { $Context.config } else { $null }
        if ($cfg -and (Get-Command -Name 'Invoke-QCProcessorWorkflowStateTelemetry' -ErrorAction SilentlyContinue)) {
            Invoke-QCProcessorWorkflowStateTelemetry -Config $cfg -Context $Context `
                -PreviousState ([string]$info.Data.stateName) -CurrentState $StateName -JobType 'QC_PREPEND' | Out-Null
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

Export-ModuleMember -Function Test-QCWorkflowConfig,Get-QCWorkflowSettings,Get-QCWorkflowDeprecationWarnings,Get-QCWorkflowStateName,Normalize-QCPrependTriggerKey,Resolve-QCWorkflowStateAfterPrepend,Resolve-QCWorkflowAssignee,Get-PWDocumentWorkflowInfo,Ensure-PWQCWorkflowAssignment,Test-QCWorkflowStateTransition,Set-PWQCWorkflowState,Set-PWQCAttributes,Invoke-QCWorkflowWriteback
