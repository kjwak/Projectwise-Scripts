# Unit checks for QC.AuditTriggers (no PW / SQL).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.AuditTriggers.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

Assert-True (Test-QCIsQcPdfDocumentName -DocumentName 'sheet-qc.pdf') 'qc pdf suffix'
Assert-True (-not (Test-QCIsQcPdfDocumentName -DocumentName 'sheet.pdf')) 'plain pdf is not qc pdf'

$cfg = @{
    auditPoller = @{
        workflowTriggers = @{
            enabled = $false
            recordStateHistory = $true
        }
    }
}
$s = Get-QCAuditWorkflowTriggerSettings -Config $cfg
Assert-Eq $s.enabled $false 'workflowTriggers.enabled respected'
Assert-Eq $s.recordStateHistory $true 'other flags still parsed'

$cfgDefault = @{ auditPoller = @{} }
$s2 = Get-QCAuditWorkflowTriggerSettings -Config $cfgDefault
Assert-Eq $s2.enabled $true 'defaults enable workflow triggers'
Assert-Eq $s2.recordFromProcessor $true 'defaults enable processor telemetry'

Import-Module (Join-Path $repoRoot 'modules\QC.WatcherOrchestration.psm1') -Force
$prependActions = Get-QCPrependAuditActions -Config $cfgDefault
Assert-True ($prependActions -contains 'DOCUMENT_ATTR') 'default prepend actions include DOCUMENT_ATTR'
Assert-True ($prependActions -contains 'DOCUMENT_STATE') 'default prepend actions include DOCUMENT_STATE'

Assert-Eq (Get-QCInitiatedWorkflowStateName -Config $cfgDefault) 'QC Initiated' 'default initiated state name'
Assert-True (Test-QCWorkflowStateIsQcInitiated -StateName 'QC Initiated' -Config $cfgDefault) 'qc initiated match'
Assert-True (-not (Test-QCWorkflowStateIsQcInitiated -StateName 'In Production' -Config $cfgDefault)) 'non-initiated state'

$cfgDbOff = @{ database = @{ enabled = $false }; auditPoller = @{ workflowTriggers = @{ enabled = $true } } }
Invoke-QCAuditWorkflowStateChangeTriggers -Config $cfgDbOff -DocumentGuid 'g1' -DocumentName 'a-qc.pdf' `
    -FolderPath 'Documents\X\CADD\Sheets' -PreviousState 'In Production' -CurrentState 'QC Received' | Out-Null
Invoke-QCAuditWorkflowAttributeChangeTriggers -Config $cfgDbOff -DocumentGuid 'g1' -DocumentName 'a.pdf' `
    -FolderPath 'Documents\X\CADD\Sheets' -FieldChanges @{ designer_email = @{ oldValue = 'a@x.com'; newValue = 'b@x.com' } } | Out-Null

$cfgAuto = @{
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            ignoreStateChangeFromAutomation = $true
            automationPwUsernames = @('srv_typsa_archivist')
            automationPwUserNumbers = @(42)
            notifyOnStateChange = $true
        }
    }
}
Assert-True (Test-QCIsAutomationPwActor -Config $cfgAuto -ChangedByUser 42) 'userno match'
Assert-True (Test-QCIsAutomationPwActor -Config $cfgAuto -ChangedByUsername 'srv_typsa_archivist') 'username match'
Assert-True (-not (Test-QCIsAutomationPwActor -Config $cfgAuto -ChangedByUser 99 -ChangedByUsername 'human')) 'non-automation'
Assert-True (-not (Test-QCShouldSuppressAuditSheetStateSync -Config $cfgAuto -DocumentName 'sheet-qc.pdf' -ChangedByUser 42)) 'allow sync on qc pdf'
Assert-True (Test-QCShouldSuppressAuditSheetStateSync -Config $cfgAuto -DocumentName 'sheet.dgn' -ChangedByUser 42) 'suppress sync echo on dgn'

$cfgAutoOff = @{ auditPoller = @{ workflowTriggers = @{ ignoreStateChangeFromAutomation = $false; automationPwUsernames = @('srv_typsa_archivist') } } }
Assert-True (-not (Test-QCIsAutomationPwActor -Config $cfgAutoOff -ChangedByUsername 'srv_typsa_archivist')) 'ignore flag off'

Write-Host 'test_audit_workflow_triggers.ps1 passed'
