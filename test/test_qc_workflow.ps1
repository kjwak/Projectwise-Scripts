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
            autoAssignWorkflow = $false
            autoSetState = $false
            autoWriteAttributes = $true
            productionStateName = 'In Production'
            receivedStateName = 'QC Received'
            correctionsInProgressStateName = 'Corrections In Progress'
            backcheckInProgressStateName = 'Backcheck In Progress'
            errorStateName = 'Error Needs Attention'
            stateAfterSuccessfulPrepend = 'Redlines Issued'
            stateAfterFailedPrepend = 'Error Needs Attention'
            attributeMap = @{
                qcActive = 'QC_Active'
                cycleId = 'QC_Cycle_ID'
                stage = 'QC_Stage'
                status = 'QC_Status'
                historyPdfPath = 'QC_History_PDF_Path'
                latestOverlayPdfPath = 'QC_Latest_Overlay_PDF_Path'
                sourceDocumentPath = 'QC_Source_Document_Path'
                automationLastRun = 'QC_Automation_Last_Run'
                automationResult = 'QC_Automation_Result'
                automationError = 'QC_Automation_Error'
            }
            stageMap = @{
                red = @{ stageValue = 'Red'; statusValue = 'Open'; optionalStateName = 'Redlines Issued' }
                green = @{ stageValue = 'Green'; statusValue = 'Pending Backcheck'; optionalStateName = 'Corrections Complete' }
                blue = @{ stageValue = 'Blue'; statusValue = 'Closed'; optionalStateName = 'Verified Closed' }
            }
        }
    }
}

function New-WorkflowContext() {
    return @{
        jobId = 'job1'
        stage = 'red'
        resultStatus = 'Succeeded'
        documentPath = 'Documents\Demo\a.pdf'
        attributes = @{
            qcActive = $true
            cycleId = 'cycle-1'
            stage = 'Red'
            status = 'Open'
            historyPdfPath = 'C:\history\a.pdf'
            latestOverlayPdfPath = 'C:\out\a.pdf'
            sourceDocumentPath = 'Documents\Demo\a.pdf'
            automationLastRun = '2026-05-07T00:00:00Z'
            automationResult = 'Succeeded'
            automationError = $null
        }
    }
}

# built-in defaults stay attribute-first and use the finalized ProjectWise state names
$defaultSettings = Get-QCWorkflowSettings -Config @{}
Assert-Eq $defaultSettings.mode 'AttributesOnly' 'AttributesOnly should remain the default workflow mode'
Assert-Eq $defaultSettings.autoSetState $false 'autoSetState should remain false by default'
Assert-Eq $defaultSettings.productionStateName 'In Production' 'Default production state name'
Assert-Eq $defaultSettings.receivedStateName 'QC Received' 'Default received state name'
Assert-Eq $defaultSettings.correctionsInProgressStateName 'Corrections In Progress' 'Default corrections in progress state name'
Assert-Eq $defaultSettings.backcheckInProgressStateName 'Backcheck In Progress' 'Default backcheck state name'
Assert-Eq $defaultSettings.errorStateName 'Error Needs Attention' 'Default error state name'
Assert-Eq $defaultSettings.stageMap.red.optionalStateName 'Redlines Issued' 'Default red optional state'
Assert-Eq $defaultSettings.stageMap.green.optionalStateName 'Corrections Complete' 'Default green optional state'
Assert-Eq $defaultSettings.stageMap.blue.optionalStateName 'Verified Closed' 'Default blue optional state'

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
function Get-PWWorkflowStateLinks { [CmdletBinding()] param() [pscustomobject]@{ FromStateName = 'QC Received'; ToStateName = 'Redlines Issued' } }
$dry = Invoke-QCWorkflowWriteback -Config (New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true) -Context (New-WorkflowContext)
Assert-True $dry.IsSuccess 'Dry-run workflow should return success'
Assert-Eq $script:pwWriteCalls 0 'Dry-run should not call PW write cmdlets'
Assert-True ((@($dry.Data.actions | Where-Object { $_.Code -eq 'QC_WORKFLOW_STATE_PLANNED' })).Count -eq 0) 'AttributesOnly dry-run should not plan state'
Assert-True ((@($dry.Data.actions | Where-Object { $_.Code -eq 'QC_WORKFLOW_ATTRIBUTES_PLANNED' })).Count -eq 1) 'Dry-run should plan attributes'

# missing workflow/state/attribute config logs warnings but does not fail by default
$missing = New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true
$missing.qcWorkflow.mode = 'NotAMode'
$missing.qcWorkflow.Remove('attributeMap')
$missing.qcWorkflow.stageMap.Remove('blue')
$v = Test-QCWorkflowConfig -Config $missing
Assert-True $v.IsSuccess 'Missing config should warn but pass when strictMode=false'
Assert-True ($v.Data.warnings.Count -ge 3) 'Missing config should include warnings'

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

# attribute map converts internal keys into configured ProjectWise attribute names correctly
$settings = Get-QCWorkflowSettings -Config (New-WorkflowConfig -Enabled:$true -Strict:$false -DryRunWriteback:$true)
$attrs = Set-PWQCAttributes -Settings $settings -Context (New-WorkflowContext) -DryRun:$true
Assert-True $attrs.IsSuccess 'Dry-run attribute mapping should succeed'
Assert-Eq $attrs.Data.attributes['QC_Cycle_ID'] 'cycle-1' 'cycleId should map to configured PW attribute'
Assert-Eq $attrs.Data.attributes['QC_Status'] 'Open' 'status should map to configured PW attribute'

# real write paths use confirmed cmdlets when not dry-run
$script:stateWrites = 0
$script:attributeWrites = 0
function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'QC Review Workflow' } }
function Get-PWWorkflowStateLinks { [CmdletBinding()] param() [pscustomobject]@{ FromStateName = 'QC Received'; ToStateName = 'Redlines Issued' } }
function Set-PWDocumentState { [CmdletBinding()] param([object[]]$InputDocuments, [string]$StateName, [switch]$ReturnBoolean) $script:stateWrites++; return $true }
function Update-PWDocumentAttributes { [CmdletBinding()] param([object[]]$InputDocuments, [hashtable]$Attributes, [switch]$ReturnBoolean) $script:attributeWrites++; return $true }
$writeCfg = New-WorkflowConfig -Enabled:$true -Strict:$true -DryRunWriteback:$false
$writeCfg.qcWorkflow.mode = 'StateAndAttributes'
$writeCfg.qcWorkflow.autoSetState = $true
$writeCfg.qcWorkflow.expectedWorkflowName = 'QC Review Workflow'
$writeSettings = Get-QCWorkflowSettings -Config $writeCfg
$doc = [pscustomobject]@{ WorkflowName = 'QC Review Workflow'; StateName = 'QC Received'; DocumentID = 1; ProjectID = 2 }
$writeContext = New-WorkflowContext
$writeContext.document = $doc
$stateWrite = Set-PWQCWorkflowState -Settings $writeSettings -Context $writeContext -StateName 'Redlines Issued' -DryRun:$false
Assert-True $stateWrite.IsSuccess 'Set-PWQCWorkflowState should succeed with confirmed Set-PWDocumentState stub'
Assert-Eq $stateWrite.Code 'QC_WORKFLOW_STATE_WRITE_SUCCESS' 'State write success code'
Assert-Eq $script:stateWrites 1 'Set-PWDocumentState should be called once'
$attrWrite = Set-PWQCAttributes -Settings $writeSettings -Context $writeContext -DryRun:$false
Assert-True $attrWrite.IsSuccess 'Set-PWQCAttributes should succeed with confirmed Update-PWDocumentAttributes stub'
Assert-Eq $attrWrite.Code 'QC_WORKFLOW_ATTRIBUTE_WRITE_SUCCESS' 'Attribute write success code'
Assert-Eq $script:attributeWrites 1 'Update-PWDocumentAttributes should be called once'

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
