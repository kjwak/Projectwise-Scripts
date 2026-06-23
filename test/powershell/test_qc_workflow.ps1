$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Workflow/QC.Workflow.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force


function Write-ValidPdf([string]$QpdfExe, [string]$Path) {
    $stdout = [System.IO.Path]::GetTempFileName()
    $stderr = [System.IO.Path]::GetTempFileName()
    try {
        $p = Start-Process -FilePath $QpdfExe -ArgumentList @('--empty','--', $Path) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        if ($p.ExitCode -ne 0) { throw (Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue) }
    } finally {
        Remove-Item -LiteralPath $stdout -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderr -Force -ErrorAction SilentlyContinue
    }
}

function New-WorkflowConfig([bool]$Enabled, [bool]$Strict, [bool]$DryRunWriteback) {
    return @{
        dryRun = $false
        qcWorkflow = @{
            enabled = $Enabled
            strictMode = $Strict
            dryRunWriteback = $DryRunWriteback
            workflowName = 'QC Review Workflow'
            expectedWorkflowName = ''
            mode = 'AttributesOnly'
            autoSetState = $false
            autoWriteAttributes = $true
            states = @{
                production = 'In Production'
                readyForQc = 'Ready for QC'
                redlinesReceived = 'Redlines Received'
                readyForVerification = 'Ready for Verification'
                qcFinalizing = 'Initiate Verification'
                complete = 'QC Complete'
                error = 'Error Needs Attention'
            }
            reviewTypes = @{
                productionQc = 'Production QC'
                peerReview = 'Peer Review'
                independentCheck = 'Independent Check'
            }
            defaultReviewType = 'Production QC'
            stateAfterSuccessfulPrepend = 'Ready for QC'
            stateAfterPrependByTrigger = @{
                initialQcPdf = 'Ready for QC'
                finalQcComplete = 'QC Complete'
            }
            stateAfterFailedPrepend = 'Error Needs Attention'
            attributeMap = @{
                qcActive = 'QC_Active'
                reviewType = 'QC_Review_Type'
                cycleId = 'QC_Cycle_ID'
                designerEmail = 'QC_Designer_Email'
                reviewerEmail = 'QC_Reviewer_Email'
                checkerEmail = 'QC_Checker_Email'
                assignedTo = 'QC_Assigned_To'
                status = 'QC_Status'
                historyPdfPath = 'QC_History_PDF_Path'
                latestOverlayPdfPath = 'QC_Latest_Overlay_PDF_Path'
                sourceDocumentPath = 'QC_Source_Document_Path'
                automationLastRun = 'QC_Automation_Last_Run'
                automationResult = 'QC_Automation_Result'
                automationError = 'QC_Automation_Error'
            }
        }
    }
}

function New-WorkflowContext() {
    return @{
        jobId = 'job1'
        resultStatus = 'Succeeded'
        documentPath = 'Documents\Demo\a.pdf'
        attributes = @{
            qcActive = $true
            reviewType = 'Production QC'
            cycleId = 'cycle-1'
            designerEmail = 'designer@example.com'
            reviewerEmail = 'reviewer@example.com'
            checkerEmail = 'checker@example.com'
            status = 'Ready for QC'
            historyPdfPath = 'C:\history\a.pdf'
            latestOverlayPdfPath = 'C:\out\a.pdf'
            sourceDocumentPath = 'Documents\Demo\a.pdf'
            automationLastRun = '2026-05-07T00:00:00Z'
            automationResult = 'Succeeded'
            automationError = $null
        }
    }
}

# built-in defaults stay attribute-first and use shared lifecycle state names
$defaultSettings = Get-QCWorkflowSettings -Config @{}
Assert-Eq $defaultSettings.mode 'AttributesOnly' 'AttributesOnly should remain the default workflow mode'
Assert-Eq $defaultSettings.autoSetState $false 'autoSetState should remain false by default'
Assert-Eq $defaultSettings.states.readyForQc 'Originated' 'Default ready for QC state'
Assert-Eq $defaultSettings.states.redlinesReceived 'Redlines Received' 'Default redlines received state'
Assert-Eq $defaultSettings.states.readyForVerification 'Ready for Verification' 'Default ready for verification state'
Assert-Eq $defaultSettings.states.qcFinalizing 'Initiate Verification' 'Default initiate verification state'
Assert-True (-not $defaultSettings.attributeMap.ContainsKey('stage')) 'Default attribute map must not include stage'

$prodCfg = @{ qcWorkflow = @{ states = @{ production = 'In Development'; readyForQc = 'Originated' } } }
Assert-Eq (Format-QCWorkflowStateName -StateName 'in development' -Config $prodCfg) 'In Development' 'Lowercase state maps to configured label'
Assert-Eq (Format-QCWorkflowStateName -StateName 'originated' -Config $prodCfg) 'Originated' 'Lowercase originated maps to configured label'
Assert-Eq (Format-QCWorkflowStateName -StateName 'redlines received' -Settings $defaultSettings) 'Redlines Received' 'Title-cases unknown multi-word states'

# review-type-based assignment
$assignReviewer = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Originated' -ReviewType 'Production QC' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignReviewer 'r@x.com' 'Production QC Originated assigns reviewer'
$assignChecker = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Ready for Verification' -ReviewType 'Independent Check' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignChecker 'c@x.com' 'Independent Check ready for verification assigns checker'
$assignProduction = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'In Development' -ReviewType 'Production QC' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignProduction 'd@x.com' 'In Development assigns designer'
$assignRedlines = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Redlines Received' -ReviewType 'Peer Review' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignRedlines 'd@x.com' 'Redlines Received assigns designer'
$assignReviewerVerify = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Ready for Verification' -ReviewType 'Peer Review' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignReviewerVerify 'r@x.com' 'Ready for Verification assigns reviewer'
$assignNone = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'QC Complete' -ReviewType 'Production QC' -ReviewerEmail 'r@x.com'
Assert-Eq $assignNone $null 'QC Complete has no active assignee'

# deprecated config warnings
$deprecatedCfg = @{
    qcWorkflow = @{
        enabled = $true
        receivedStateName = 'QC Received'
        stageMap = @{ red = @{ stageValue = 'Red' } }
        attributeMap = @{ qcActive = 'QC_Active'; stage = 'QC_Stage' }
    }
}
$depWarnings = @(Get-QCWorkflowDeprecationWarnings -RawWorkflowConfig $deprecatedCfg.qcWorkflow)
Assert-True ($depWarnings.Count -ge 2) 'Deprecated keys should produce warnings'

# Get-PWWorkflowStateLinks must receive WorkflowName (no unfiltered call — avoids interactive prompt)
function Get-PWWorkflowStateLinks {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$WorkflowName)
    if ([string]::IsNullOrWhiteSpace($WorkflowName)) { throw 'WorkflowName required' }
    return [pscustomobject]@{ FromStateName = 'In Production'; ToStateName = 'Ready for QC' }
}
$mandatorySettings = Get-QCWorkflowSettings -Config (New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true)
$mandatorySettings.expectedWorkflowName = 'TYPSA QC'
$tr = Test-QCWorkflowStateTransition -Settings $mandatorySettings -CurrentStateName 'In Production' -TargetStateName 'Ready for QC' -WorkflowName 'TYPSA QC'
Assert-True $tr.IsSuccess 'Transition test should succeed with mandatory WorkflowName stub'
Remove-Item function:\Get-PWWorkflowStateLinks -ErrorAction SilentlyContinue

# workflow config disabled means no writeback occurs
$disabled = Invoke-QCWorkflowWriteback -Config (New-WorkflowConfig -Enabled:$false -Strict:$false -DryRunWriteback:$true) -Context (New-WorkflowContext)
Assert-True $disabled.IsSuccess 'Disabled workflow should return success'
Assert-Eq $disabled.Code 'QC_WORKFLOW_DISABLED' 'Disabled workflow code'
Assert-Eq $disabled.Data.actions.Count 0 'Disabled workflow should not plan actions'

# dry-run returns planned changes and does not call PW write cmdlets
$script:pwWriteCalls = 0
function Set-PWDocumentState { $script:pwWriteCalls++; throw 'Should not be called during dry-run' }
function Update-PWDocumentAttributes { $script:pwWriteCalls++; throw 'Should not be called during dry-run' }
function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'QC Review Workflow' } }
function Get-PWWorkflowStateLinks { [CmdletBinding()] param() [pscustomobject]@{ FromStateName = 'Ready for QC'; ToStateName = 'Redlines Received' } }
$dry = Invoke-QCWorkflowWriteback -Config (New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true) -Context (New-WorkflowContext)
Assert-True $dry.IsSuccess 'Dry-run workflow should return success'
Assert-Eq $script:pwWriteCalls 0 'Dry-run should not call PW write cmdlets'
Assert-True ((@($dry.Data.actions | Where-Object { $_.Code -eq 'QC_WORKFLOW_STATE_PLANNED' })).Count -eq 0) 'AttributesOnly dry-run should not plan state'
Assert-True ((@($dry.Data.actions | Where-Object { $_.Code -eq 'QC_WORKFLOW_ATTRIBUTES_PLANNED' })).Count -eq 1) 'Dry-run should plan attributes'

# missing workflow config logs warnings but does not fail by default
$missing = New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true
$missing.qcWorkflow.mode = 'NotAMode'
$missing.qcWorkflow.Remove('attributeMap')
$v = Test-QCWorkflowConfig -Config $missing
Assert-True $v.IsSuccess 'Missing config should warn but pass when strictMode=false'
Assert-True ($v.Data.warnings.Count -ge 1) 'Missing config should include warnings'

# strictMode converts missing configuration into failure
$strict = New-WorkflowConfig -Enabled:$true -Strict:$true -DryRunWriteback:$true
$strict.qcWorkflow.mode = 'NotAMode'
$strictRes = Test-QCWorkflowConfig -Config $strict
Assert-True (-not $strictRes.IsSuccess) 'Strict invalid workflow mode should fail validation'
Assert-Eq $strictRes.Code 'QC_WORKFLOW_CONFIG_STRICT_FAILURE' 'Strict validation code'

# workflow writeback failure does not fail unless strictMode = true
$nonStrictNoCmdlets = New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$false
Remove-Item function:\Set-PWDocumentState -ErrorAction SilentlyContinue
Remove-Item function:\Update-PWDocumentAttributes -ErrorAction SilentlyContinue
Remove-Item function:\Get-PWWorkflows -ErrorAction SilentlyContinue
Remove-Item function:\Get-PWWorkflowStateLinks -ErrorAction SilentlyContinue
$nonStrict = Invoke-QCWorkflowWriteback -Config $nonStrictNoCmdlets -Context (New-WorkflowContext)
Assert-True $nonStrict.IsSuccess 'Missing PW write cmdlets should be non-fatal when strictMode=false'
Assert-True ($nonStrict.Data.warnings.Count -ge 1) 'Missing PW write cmdlets should return warnings'

$strictNoAttrs = New-WorkflowConfig -Enabled:$true -Strict:$true -DryRunWriteback:$true
$strictNoAttrs.qcWorkflow.Remove('attributeMap')
$strictFail = Invoke-QCWorkflowWriteback -Config $strictNoAttrs -Context (New-WorkflowContext)
Assert-True (-not $strictFail.IsSuccess) 'Strict validation failure should fail workflow writeback'

# attribute map converts internal keys into configured ProjectWise attribute names; no QC_Stage
$settings = Get-QCWorkflowSettings -Config (New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true)
$attrs = Set-PWQCAttributes -Settings $settings -Context (New-WorkflowContext) -DryRun:$true
Assert-True $attrs.IsSuccess 'Dry-run attribute mapping should succeed'
Assert-Eq $attrs.Data.attributes['QC_Cycle_ID'] 'cycle-1' 'cycleId should map to configured PW attribute'
Assert-Eq $attrs.Data.attributes['QC_Review_Type'] 'Production QC' 'reviewType should map to configured PW attribute'
Assert-True (-not $attrs.Data.attributes.ContainsKey('QC_Stage')) 'QC_Stage must not be written'

# Phase 1: automation-only attributes stay in context but are excluded from PW writeback
$phase1ExcludedPw = @(
    'QC_Active'
    'QC_Last_Action_By'
    'QC_Last_Action_Date'
    'QC_History_PDF_Path'
    'QC_Latest_Overlay_PDF_Path'
    'QC_Source_Document_Path'
    'QC_Automation_Last_Run'
    'QC_Automation_Result'
    'QC_Automation_Error'
)
foreach ($pwAttr in $phase1ExcludedPw) {
    Assert-True (-not $attrs.Data.attributes.ContainsKey($pwAttr)) "Phase 1 excluded PW attribute must not be written: $pwAttr"
}
Assert-Eq $attrs.Data.internalAttributes['historyPdfPath'] 'C:\history\a.pdf' 'Full context should retain historyPdfPath internally'
$expectedSkipped = @('qcActive','historyPdfPath','latestOverlayPdfPath','sourceDocumentPath','automationLastRun','automationResult','automationError')
foreach ($key in $expectedSkipped) {
    Assert-True ($attrs.Data.skippedWritebackKeys -contains $key) "skippedWritebackKeys should include $key"
}
$defaults = Get-QCWorkflowAttributeWritebackExcludeDefaults
Assert-True ($defaults -contains 'automationResult') 'Default exclude list should include automationResult'

$legacyCfg = New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true
$legacyCfg.qcWorkflow.attributeWritebackExcludeDisabled = $true
$legacySettings = Get-QCWorkflowSettings -Config $legacyCfg
$legacyAttrs = Set-PWQCAttributes -Settings $legacySettings -Context (New-WorkflowContext) -DryRun:$true
Assert-True $legacyAttrs.Data.attributes.ContainsKey('QC_History_PDF_Path') 'Legacy mode should write historyPdfPath to PW'
Assert-Eq $legacyAttrs.Data.skippedWritebackKeys.Count 0 'Legacy mode should not skip writeback keys'

# prepend trigger resolves target state from module defaults (TYPSA-aligned; $defaultSettings from line 105)
Assert-Eq (Resolve-QCWorkflowStateAfterPrepend -Settings $defaultSettings -Context @{ prependTrigger = 'initialQcPdf' }) 'Originated' 'initialQcPdf -> Originated'
Assert-Eq (Resolve-QCWorkflowStateAfterPrepend -Settings $defaultSettings -Context @{ prependTrigger = 'finalQcComplete' }) 'Ready for Verification' 'finalQcComplete -> Ready for Verification'
Assert-Eq (Resolve-QCWorkflowStateAfterPrepend -Settings $defaultSettings -Context @{}) 'Ready for QC' 'unknown trigger falls back to stateAfterSuccessfulPrepend module default'

# state write uses explicit targetState when provided
$global:qcWorkflowTestStateWrites = 0
$global:qcWorkflowTestAttributeWrites = 0
$writeCfg = New-WorkflowConfig -Enabled:$true -Strict:$true -DryRunWriteback:$false
$writeCfg.qcWorkflow.mode = 'StateAndAttributes'
$writeCfg.qcWorkflow.autoSetState = $true
$writeCfg.qcWorkflow.expectedWorkflowName = 'QC Review Workflow'
$writeSettings = Get-QCWorkflowSettings -Config $writeCfg
$doc = [pscustomobject]@{ WorkflowName = 'QC Review Workflow'; StateName = 'Ready for QC'; DocumentID = 1; ProjectID = 2; DocumentGUID = 'qc-workflow-test-doc-1' }
$writeContext = New-WorkflowContext
$writeContext.document = $doc
$writeContext.config = $writeCfg
$writeContext.skipSiblingStateSync = $true
$global:qcWorkflowTestWriteSettings = $writeSettings
$global:qcWorkflowTestWriteContext = $writeContext
Remove-QCModuleFlatShims -ModuleName 'QC.Workflow.psm1'
InModuleScope -ModuleName QC.Workflow {
    function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'QC Review Workflow' } }
    function Get-PWWorkflowStateLinks { [CmdletBinding()] param() [pscustomobject]@{ FromStateName = 'Ready for QC'; ToStateName = 'Redlines Received' } }
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName, $IsLaneAuthority, $WriteScope)
        $global:qcWorkflowTestStateWrites++
        return @{ applied = $true; verified = $true; readBackState = $TargetState; targetState = $TargetState }
    }
    $stateWrite = Set-PWQCWorkflowState -Settings $global:qcWorkflowTestWriteSettings -Context $global:qcWorkflowTestWriteContext -StateName 'Redlines Received' -DryRun:$false
    if (-not $stateWrite.IsSuccess) { throw 'ASSERT FAILED: Set-PWQCWorkflowState should succeed with confirmed Set-PWDocumentState stub' }
    if ($stateWrite.Code -ne 'QC_WORKFLOW_STATE_WRITE_SUCCESS') { throw "ASSERT FAILED: State write success code`nExpected: QC_WORKFLOW_STATE_WRITE_SUCCESS`nActual:   $($stateWrite.Code)" }
}
Assert-Eq $global:qcWorkflowTestStateWrites 1 'Set-PWDocumentState should be called once'
function Update-PWDocumentAttributes { [CmdletBinding()] param([object[]]$InputDocuments, [hashtable]$Attributes, [switch]$ReturnBoolean) $global:qcWorkflowTestAttributeWrites++; return $true }
$attrWrite = Set-PWQCAttributes -Settings $writeSettings -Context $writeContext -DryRun:$false
Assert-True $attrWrite.IsSuccess 'Set-PWQCAttributes should succeed with confirmed Update-PWDocumentAttributes stub'
Assert-Eq $attrWrite.Code 'QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS' 'Attribute write success code'
Assert-Eq $global:qcWorkflowTestAttributeWrites 1 'Update-PWDocumentAttributes should be called once'

# empty Get-PWWorkflowStateLinks: fall back to qcWorkflow.states when strictMode=false
$global:qcWorkflowTestStateWrites = 0
$fallbackCfg = New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$false
$fallbackCfg.qcWorkflow.mode = 'StateAndAttributes'
$fallbackCfg.qcWorkflow.autoSetState = $true
$fallbackCfg.qcWorkflow.expectedWorkflowName = 'TYPSA QC'
$fallbackSettings = Get-QCWorkflowSettings -Config $fallbackCfg
$fallbackDoc = [pscustomobject]@{ WorkflowName = 'TYPSA QC'; StateName = 'In Production'; DocumentID = 1; ProjectID = 2; DocumentGUID = 'qc-workflow-test-doc-2' }
$fallbackCtx = New-WorkflowContext
$fallbackCtx.document = $fallbackDoc
$fallbackCtx.config = $fallbackCfg
$fallbackCtx.skipSiblingStateSync = $true
$global:qcWorkflowTestFallbackSettings = $fallbackSettings
$global:qcWorkflowTestFallbackCtx = $fallbackCtx
Remove-QCModuleFlatShims -ModuleName 'QC.Workflow.psm1'
InModuleScope -ModuleName QC.Workflow {
    function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'TYPSA QC' } }
    function Get-PWWorkflowStateLinks { [CmdletBinding()] param() @() }
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName, $IsLaneAuthority, $WriteScope)
        $global:qcWorkflowTestStateWrites++
        return @{ applied = $true; verified = $true; readBackState = $TargetState; targetState = $TargetState }
    }
    $fallbackWrite = Set-PWQCWorkflowState -Settings $global:qcWorkflowTestFallbackSettings -Context $global:qcWorkflowTestFallbackCtx -StateName 'Ready for QC' -DryRun:$false
    if (-not $fallbackWrite.IsSuccess) { throw 'ASSERT FAILED: Set-PWQCWorkflowState should succeed via qcWorkflow.states fallback when links are empty' }
    if ($fallbackWrite.Code -ne 'QC_WORKFLOW_STATE_WRITE_SUCCESS') { throw "ASSERT FAILED: Configured fallback should still write state`nExpected: QC_WORKFLOW_STATE_WRITE_SUCCESS`nActual:   $($fallbackWrite.Code)" }
}
Assert-Eq $global:qcWorkflowTestStateWrites 1 'Set-PWDocumentState should run once on configured fallback'

# successful QC_PREPEND calls Invoke-QCWorkflowWriteback (disabled default returns attached workflow result)
$tmp = Join-Path $env:TEMP ("qc-workflow-prepend-" + ([guid]::NewGuid().ToString('N')))
try {
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $src = Join-Path $tmp 'a.pdf'
    $qpdf = Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe'
    Write-ValidPdf -QpdfExe $qpdf -Path $src
    $cfg = @{
        dryRun = $false
        fileOpThrottleMs = 0
        qcWorkflow = @{ enabled = $false }
        qcPrepend = @{ historyRoot = (Join-Path $tmp 'history'); tempRoot = (Join-Path $tmp 'temp'); outputRoot = (Join-Path $tmp 'output'); enableOverlay = $false; qpdfExePath = $qpdf }
    }
    $job = @{ id='prepend'; type='QC_PREPEND'; sourcePath=$src; sourceName='a.pdf'; sourceFolder=$tmp; metadata=@{ qcProcessType = 'production' } }
    $pre = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $pre.IsSuccess 'QC_PREPEND should succeed'
    Assert-Eq $pre.Code 'QC_PREPEND_OK' 'QC_PREPEND success code'
    Assert-True ($null -ne $pre.Data.workflowWriteback) 'QC_PREPEND should attach workflow writeback result'
    Assert-Eq $pre.Data.workflowWriteback.Code 'QC_WORKFLOW_DISABLED' 'Disabled workflow writeback should be called and attached'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'test_qc_workflow.ps1 passed'
