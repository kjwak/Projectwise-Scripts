import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
APPSETTINGS = REPO_ROOT / "appsettings.json"
WORKFLOW = REPO_ROOT / "modules" / "QC.Workflow.psm1"
REPORTING_DOC = REPO_ROOT / "docs" / "qc-reporting.md"


def test_appsettings_qc_workflow_defaults_remain_disabled_and_attribute_first():
    config = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))
    workflow = config["qcWorkflow"]

    assert workflow["enabled"] is False
    assert workflow["mode"] == "AttributesOnly"
    assert workflow["autoSetState"] is False
    assert workflow["dryRunWriteback"] is True
    assert workflow["strictMode"] is False


def test_appsettings_contains_final_projectwise_state_names():
    workflow = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))["qcWorkflow"]

    assert workflow["productionStateName"] == "In Production"
    assert workflow["receivedStateName"] == "QC Received"
    assert workflow["correctionsInProgressStateName"] == "Corrections In Progress"
    assert workflow["backcheckInProgressStateName"] == "Backcheck In Progress"
    assert workflow["errorStateName"] == "Error Needs Attention"
    assert workflow["defaultStateAfterPrepend"] == "QC Received"
    assert workflow["stateAfterSuccessfulPrepend"] == "Redlines Issued"
    assert workflow["stateAfterFailedPrepend"] == "Error Needs Attention"


def test_appsettings_stage_map_matches_initial_workflow_model():
    stage_map = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))["qcWorkflow"]["stageMap"]

    assert stage_map["red"] == {
        "stageValue": "Red",
        "statusValue": "Open",
        "optionalStateName": "Redlines Issued",
    }
    assert stage_map["green"] == {
        "stageValue": "Green",
        "statusValue": "Pending Backcheck",
        "optionalStateName": "Corrections Complete",
    }
    assert stage_map["blue"] == {
        "stageValue": "Blue",
        "statusValue": "Closed",
        "optionalStateName": "Verified Closed",
    }


def test_workflow_module_fallback_defaults_use_final_state_names():
    text = WORKFLOW.read_text(encoding="utf-8")

    for expected in [
        "productionStateName = 'In Production'",
        "receivedStateName = 'QC Received'",
        "correctionsInProgressStateName = 'Corrections In Progress'",
        "backcheckInProgressStateName = 'Backcheck In Progress'",
        "errorStateName = 'Error Needs Attention'",
        "red = @{ stageValue = 'Red'; statusValue = 'Open'; optionalStateName = 'Redlines Issued' }",
        "blue = @{ stageValue = 'Blue'; statusValue = 'Closed'; optionalStateName = 'Verified Closed' }",
    ]:
        assert expected in text

    assert "optionalStateName = 'Corrections Required'" not in text
    assert "optionalStateName = 'QC Verified'" not in text
    assert "Error / Needs Attention" not in text


def test_reporting_docs_include_new_metric_names():
    text = REPORTING_DOC.read_text(encoding="utf-8")

    for metric in [
        "inProductionCount",
        "qcReceivedCount",
        "redlinesIssuedCount",
        "correctionsInProgressCount",
        "correctionsCompleteCount",
        "backcheckInProgressCount",
        "verifiedClosedCount",
        "errorNeedsAttentionCount",
        "staleOpenQcCount",
    ]:
        assert metric in text
