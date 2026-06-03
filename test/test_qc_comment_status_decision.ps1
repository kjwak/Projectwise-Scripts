# Pure decision engine tests (no PW/DB/file I/O).
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/QC.CommentStatusDecision.psm1" -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$config = @{
    qcCommentSync = @{
        reviewerAuthorPatterns = @('(?i)reviewer')
        statusMappings = @{
            open = @('Open', 'None', '', 'Unknown')
            resolved = @('Completed', 'Accepted', 'Checked')
            closed = @('Closed', 'Cancelled')
        }
        targetStates = @{
            redlinesReceived = 'Redlines Received'
            correctionsReceived = 'Corrections Received'
            qcFinalizing = 'QC Finalizing'
            completed = 'QC Complete'
            error = 'Error Needs Attention'
        }
    }
}

$openAnnot = @{
    annotation_id = '1'; author = 'QC Reviewer'; status = 'Open'
}
$resolvedAnnot = @{
    annotation_id = '2'; author = 'QC Reviewer'; status = 'Completed'
}
$closedAnnot = @{
    annotation_id = '3'; author = 'QC Reviewer'; status = 'Closed'
}

$d1 = Resolve-QCCommentWorkflowState -Annotations @($openAnnot) -Config $config -ParserStatus 'ok'
Assert-Eq $d1.targetState 'Redlines Received' 'Open reviewer comment -> redlines received'
Assert-Eq $d1.decisionCode 'REDLINES_RECEIVED' 'Decision code redlines received'

$d2 = Resolve-QCCommentWorkflowState -Annotations @($resolvedAnnot) -Config $config -ParserStatus 'ok'
Assert-Eq $d2.targetState 'Corrections Received' 'Resolved -> corrections received'
Assert-Eq $d2.decisionCode 'CORRECTIONS_RECEIVED' 'Decision code corrections received'

$d3 = Resolve-QCCommentWorkflowState -Annotations @($closedAnnot) -Config $config -ParserStatus 'ok'
Assert-Eq $d3.targetState 'QC Finalizing' 'Closed -> QC Finalizing'
Assert-Eq $d3.decisionCode 'QC_FINALIZING' 'Decision code QC Finalizing'

$d4 = Resolve-QCCommentWorkflowState -Annotations @() -Config $config -ParserStatus 'error'
Assert-Eq $d4.targetState 'Error Needs Attention' 'Parse error -> error state'
Assert-Eq $d4.decisionCode 'PARSE_ERROR' 'Parse error code'

Write-Host 'OK test_qc_comment_status_decision.ps1'
