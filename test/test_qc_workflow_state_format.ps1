$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.Workflow.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }

$cfg = @{
    qcWorkflow = @{
        states = @{
            production = 'In Development'
            readyForQc = 'Originated'
        }
    }
}

Assert-Eq (Format-QCWorkflowStateName -StateName 'in development' -Config $cfg) 'In Development' 'configured production state'
Assert-Eq (Format-QCWorkflowStateName -StateName 'ORIGINATED' -Config $cfg) 'Originated' 'configured originated state'
Assert-Eq (Format-QCWorkflowStateName -StateName 'initiate origination' -Config $cfg) 'Initiate Origination' 'title-case fallback'

Write-Host 'test_qc_workflow_state_format: OK' -ForegroundColor Green
