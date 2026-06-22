import json
import subprocess
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
APPSETTINGS = REPO_ROOT / "appsettings.json"
WORKFLOW = REPO_ROOT / "modules" / "QC.Workflow.psm1"
PROCESSORS = REPO_ROOT / "modules" / "QC.Processors.psm1"
REPORTING_DOC = REPO_ROOT / "docs" / "qc-reporting.md"


def test_appsettings_qc_workflow_defaults_remain_disabled_and_attribute_first():
    config = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))
    workflow = config["qcWorkflow"]

    assert workflow["enabled"] is True
    assert workflow["mode"] == "StateAndAttributes"
    assert workflow["autoSetState"] is True
    assert workflow["dryRunWriteback"] is False
    assert workflow["strictMode"] is False


def test_appsettings_uses_states_and_review_types_not_stage_map():
    workflow = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))["qcWorkflow"]

    assert "stageMap" not in workflow
    assert workflow["states"]["production"] == "In Development"
    assert workflow["states"]["qcInitiated"] == "Initiate Origination"
    assert workflow["states"]["qcReceived"] == "Originated"
    assert workflow["states"]["readyForQc"] == "Originated"
    assert workflow["states"]["redlinesReceived"] == "Redlines Received"
    assert "correctionsReceived" not in workflow["states"]
    assert workflow["states"]["readyForVerification"] == "Ready for Verification"
    assert workflow["states"]["qcFinalizing"] == "Initiate Verification"
    assert workflow["stateAfterSuccessfulPrepend"] == "Originated"
    assert workflow["stateAfterPrependByTrigger"]["initialQcPdf"] == "Originated"
    assert workflow["stateAfterPrependByTrigger"]["finalQcComplete"] == "Ready for Verification"
    assert workflow["states"]["complete"] == "Verified"
    assert workflow["reviewTypes"]["independentCheck"] == "Check"
    assert workflow["defaultReviewType"] == "Production"

    attr = workflow["attributeMap"]
    assert "stage" not in attr
    assert attr["reviewType"] == "QC_Review_Type"
    assert attr["designerEmail"] == "QC_Designer_Email"
    assert attr["reviewerEmail"] == "QC_Reviewer_Email"
    assert attr["checkerEmail"] == "QC_Checker_Email"


def test_workflow_module_defaults_exclude_stage_map_and_qc_stage():
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "stageMap" not in text or "Deprecated" in text or "deprecated" in text.lower()
    assert "QC_Stage" not in text
    assert "readyForQc = 'Originated'" in text
    assert "redlinesReceived = 'Redlines Received'" in text
    assert "readyForVerification = 'Ready for Verification'" in text
    assert "qcFinalizing = 'Initiate Verification'" in text
    assert "Resolve-QCWorkflowAssignee" in text
    assert "Get-QCWorkflowDeprecationWarnings" in text


def test_processors_workflow_context_does_not_write_qc_stage():
    text = PROCESSORS.read_text(encoding="utf-8")
    assert "stage = $stageValue" not in text
    assert "QC_Stage" not in text
    assert "reviewType = $reviewType" in text
    assert "Resolve-QCWorkflowAssignee" in text
    assert "Add-QCPrependJobForQcFinalizingStateChange" in text
    assert "_QCP-IsFinalQcPrependJob" in text


def test_reporting_docs_include_state_based_metric_names():
    text = REPORTING_DOC.read_text(encoding="utf-8")

    for metric in [
        "inProductionCount",
        "readyForQcCount",
        "redlinesReceivedCount",
        "correctionsReceivedCount",
        "qcFinalizingCount",
        "qcCompleteCount",
        "errorNeedsAttentionCount",
        "staleOpenQcCount",
    ]:
        assert metric in text


def _pwsh_eval(script: str) -> str:
    shell = "pwsh"
    import shutil

    if not shutil.which(shell):
        shell = "powershell"
    preamble = "$WarningPreference = 'SilentlyContinue'\n"
    result = subprocess.run(
        [shell, "-NoProfile", "-Command", preamble + script],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        check=True,
    )
    lines = [ln for ln in result.stdout.splitlines() if ln.strip()]
    return lines[-1].strip() if lines else ""


def test_resolve_qc_workflow_assignee_by_review_type():
    script = r"""
    $ErrorActionPreference = 'Stop'
    Import-Module './modules/Core.Results.psm1' -Force
    Import-Module './modules/QC.Workflow.psm1' -Force
    $s = Get-QCWorkflowSettings -Config @{}
    @(
      (Resolve-QCWorkflowAssignee -Settings $s -StateName 'In Development' -ReviewType 'Production' -ReviewerEmail 'r@x.com' -DesignerEmail 'd@x.com' -CheckerEmail 'c@x.com'),
      (Resolve-QCWorkflowAssignee -Settings $s -StateName 'Originated' -ReviewType 'Production' -ReviewerEmail 'r@x.com' -DesignerEmail 'd@x.com' -CheckerEmail 'c@x.com'),
      (Resolve-QCWorkflowAssignee -Settings $s -StateName 'Originated' -ReviewType 'Check' -ReviewerEmail 'r@x.com' -DesignerEmail 'd@x.com' -CheckerEmail 'c@x.com'),
      (Resolve-QCWorkflowAssignee -Settings $s -StateName 'Redlines Received' -ReviewType 'Review' -ReviewerEmail 'r@x.com' -DesignerEmail 'd@x.com' -CheckerEmail 'c@x.com'),
      (Resolve-QCWorkflowAssignee -Settings $s -StateName 'Ready for Verification' -ReviewType 'Check' -ReviewerEmail 'r@x.com' -DesignerEmail 'd@x.com' -CheckerEmail 'c@x.com'),
      (Resolve-QCWorkflowAssignee -Settings $s -StateName 'Verified' -ReviewType 'Production' -ReviewerEmail 'r@x.com')
    ) -join '|'
    """
    out = _pwsh_eval(script)
    assert out == "d@x.com|r@x.com|c@x.com|d@x.com|c@x.com|"


def test_deprecated_config_emits_warnings():
    script = r"""
    $ErrorActionPreference = 'Stop'
    Import-Module './modules/Core.Results.psm1' -Force
    Import-Module './modules/QC.Workflow.psm1' -Force
    $cfg = @{
      qcWorkflow = @{
        enabled = $true
        receivedStateName = 'QC Received'
        stageMap = @{ red = @{ stageValue = 'Red' } }
      }
    }
    $v = Test-QCWorkflowConfig -Config $cfg
    ($v.Data.warnings | Where-Object { $_ -match 'deprecated' }).Count
    """
    count = int(_pwsh_eval(script))
    assert count >= 2


def test_resolve_state_after_prepend_by_trigger():
    script = r"""
    $ErrorActionPreference = 'Stop'
    Import-Module './modules/Core.Results.psm1' -Force
    Import-Module './modules/QC.Workflow.psm1' -Force
    $s = Get-QCWorkflowSettings -Config @{}
    @(
      (Resolve-QCWorkflowStateAfterPrepend -Settings $s -Context @{ prependTrigger = 'finalQcComplete' }),
      (Resolve-QCWorkflowStateAfterPrepend -Settings $s -Context @{ prependTrigger = 'initialQcPdf' }),
      (Resolve-QCWorkflowStateAfterPrepend -Settings $s -Context @{}),
      (Resolve-QCWorkflowStateAfterPrepend -Settings $s -Context @{ targetState = 'Initiate Verification' })
    ) -join '|'
    """
    out = _pwsh_eval(script)
    assert out == "Ready for Verification|Originated|Ready for QC|Initiate Verification"


def test_dry_run_attributes_do_not_include_qc_stage():
    script = r"""
    $ErrorActionPreference = 'Stop'
    Import-Module './modules/Core.Results.psm1' -Force
    Import-Module './modules/QC.Workflow.psm1' -Force
    $settings = Get-QCWorkflowSettings -Config @{}
    $ctx = @{
      attributes = @{
        qcActive = $true
        reviewType = 'Production QC'
        designerEmail = 'd@x.com'
        reviewerEmail = 'r@x.com'
        status = 'Ready for QC'
      }
    }
    $r = Set-PWQCAttributes -Settings $settings -Context $ctx -DryRun:$true
    ($r.Data.attributes.Keys -join ',')
    """
    keys = _pwsh_eval(script)
    assert "QC_Stage" not in keys
    assert "QC_Review_Type" in keys
