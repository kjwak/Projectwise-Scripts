$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/QC.Workflow.psm1" -Force
Import-Module "$repoRoot/modules/QC.Processors.psm1" -Force


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
                correctionsReceived = 'Corrections Received'
                qcFinalizing = 'QC Finalizing'
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
Assert-Eq $defaultSettings.states.readyForQc 'Ready for QC' 'Default ready for QC state'
Assert-Eq $defaultSettings.states.redlinesReceived 'Redlines Received' 'Default redlines received state'
Assert-Eq $defaultSettings.states.correctionsReceived 'Corrections Received' 'Default corrections received state'
Assert-Eq $defaultSettings.states.qcFinalizing 'QC Finalizing' 'Default QC Finalizing state'
Assert-True (-not $defaultSettings.attributeMap.ContainsKey('stage')) 'Default attribute map must not include stage'

# review-type-based assignment
$assignReviewer = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Ready for QC' -ReviewType 'Production QC' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignReviewer 'r@x.com' 'Production QC Ready for QC assigns reviewer'
$assignChecker = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Corrections Received' -ReviewType 'Independent Check' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignChecker 'c@x.com' 'Independent Check corrections received assigns checker'
$assignProduction = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'In Production' -ReviewType 'Production QC' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignProduction 'd@x.com' 'In Production assigns designer'
$assignRedlines = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Redlines Received' -ReviewType 'Peer Review' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignRedlines 'd@x.com' 'Redlines Received assigns designer'
$assignReviewerVerify = Resolve-QCWorkflowAssignee -Settings $defaultSettings -StateName 'Corrections Received' -ReviewType 'Peer Review' `
    -DesignerEmail 'd@x.com' -ReviewerEmail 'r@x.com' -CheckerEmail 'c@x.com'
Assert-Eq $assignReviewerVerify 'r@x.com' 'Corrections Received assigns reviewer'
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

# prepend trigger selects target state when autoSetState runs
$triggerSettings = Get-QCWorkflowSettings -Config $writeCfg
Assert-Eq (Resolve-QCWorkflowStateAfterPrepend -Settings $triggerSettings -Context @{ prependTrigger = 'initialQcPdf' }) 'Ready for QC' 'initialQcPdf -> Ready for QC'
Assert-Eq (Resolve-QCWorkflowStateAfterPrepend -Settings $triggerSettings -Context @{ prependTrigger = 'finalQcComplete' }) 'QC Complete' 'finalQcComplete -> QC Complete'
Assert-Eq (Resolve-QCWorkflowStateAfterPrepend -Settings $triggerSettings -Context @{}) 'Ready for QC' 'unknown trigger falls back to stateAfterSuccessfulPrepend'

# state write uses explicit targetState when provided
$script:stateWrites = 0
$script:attributeWrites = 0
function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'QC Review Workflow' } }
function Get-PWWorkflowStateLinks { [CmdletBinding()] param() [pscustomobject]@{ FromStateName = 'Ready for QC'; ToStateName = 'Redlines Received' } }
function Set-PWDocumentState { [CmdletBinding()] param([object[]]$InputDocuments, [string]$StateName, [switch]$ReturnBoolean) $script:stateWrites++; return $true }
function Update-PWDocumentAttributes { [CmdletBinding()] param([object[]]$InputDocuments, [hashtable]$Attributes, [switch]$ReturnBoolean) $script:attributeWrites++; return $true }
$writeCfg = New-WorkflowConfig -Enabled:$true -Strict:$true -DryRunWriteback:$false
$writeCfg.qcWorkflow.mode = 'StateAndAttributes'
$writeCfg.qcWorkflow.autoSetState = $true
$writeCfg.qcWorkflow.expectedWorkflowName = 'QC Review Workflow'
$writeSettings = Get-QCWorkflowSettings -Config $writeCfg
$doc = [pscustomobject]@{ WorkflowName = 'QC Review Workflow'; StateName = 'Ready for QC'; DocumentID = 1; ProjectID = 2 }
$writeContext = New-WorkflowContext
$writeContext.document = $doc
$stateWrite = Set-PWQCWorkflowState -Settings $writeSettings -Context $writeContext -StateName 'Redlines Received' -DryRun:$false
Assert-True $stateWrite.IsSuccess 'Set-PWQCWorkflowState should succeed with confirmed Set-PWDocumentState stub'
Assert-Eq $stateWrite.Code 'QC_WORKFLOW_STATE_WRITE_SUCCESS' 'State write success code'
Assert-Eq $script:stateWrites 1 'Set-PWDocumentState should be called once'
$attrWrite = Set-PWQCAttributes -Settings $writeSettings -Context $writeContext -DryRun:$false
Assert-True $attrWrite.IsSuccess 'Set-PWQCAttributes should succeed with confirmed Update-PWDocumentAttributes stub'
Assert-Eq $attrWrite.Code 'QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS' 'Attribute write success code'
Assert-Eq $script:attributeWrites 1 'Update-PWDocumentAttributes should be called once'

# empty Get-PWWorkflowStateLinks: fall back to qcWorkflow.states when strictMode=false
$script:stateWrites = 0
function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'TYPSA QC' } }
function Get-PWWorkflowStateLinks { [CmdletBinding()] param() @() }
function Set-PWDocumentState { [CmdletBinding()] param([object[]]$InputDocuments, [string]$StateName, [switch]$ReturnBoolean) $script:stateWrites++; return $true }
$fallbackCfg = New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$false
$fallbackCfg.qcWorkflow.mode = 'StateAndAttributes'
$fallbackCfg.qcWorkflow.autoSetState = $true
$fallbackCfg.qcWorkflow.expectedWorkflowName = 'TYPSA QC'
$fallbackSettings = Get-QCWorkflowSettings -Config $fallbackCfg
$fallbackDoc = [pscustomobject]@{ WorkflowName = 'TYPSA QC'; StateName = 'In Production'; DocumentID = 1; ProjectID = 2 }
$fallbackCtx = New-WorkflowContext
$fallbackCtx.document = $fallbackDoc
$fallbackWrite = Set-PWQCWorkflowState -Settings $fallbackSettings -Context $fallbackCtx -StateName 'Ready for QC' -DryRun:$false
Assert-True $fallbackWrite.IsSuccess 'Set-PWQCWorkflowState should succeed via qcWorkflow.states fallback when links are empty'
Assert-Eq $fallbackWrite.Code 'QC_WORKFLOW_STATE_WRITE_SUCCESS' 'Configured fallback should still write state'
Assert-Eq $script:stateWrites 1 'Set-PWDocumentState should run once on configured fallback'

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
    $job = @{ id='prepend'; type='QC_PREPEND'; sourcePath=$src; sourceName='a.pdf'; sourceFolder=$tmp; metadata=@{} }
    $pre = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $pre.IsSuccess 'QC_PREPEND should succeed'
    Assert-Eq $pre.Code 'QC_PREPEND_OK' 'QC_PREPEND success code'
    Assert-True ($null -ne $pre.Data.workflowWriteback) 'QC_PREPEND should attach workflow writeback result'
    Assert-Eq $pre.Data.workflowWriteback.Code 'QC_WORKFLOW_DISABLED' 'Disabled workflow writeback should be called and attached'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'test_qc_workflow.ps1 passed'
