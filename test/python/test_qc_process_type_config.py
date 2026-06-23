"""Static validation for QCProcess config shape."""
from pathlib import Path
import json

ROOT = Path(__file__).resolve().parents[1]


def test_qc_process_config_shape() -> None:
    appsettings = json.loads((ROOT / "appsettings.json").read_text(encoding="utf-8-sig"))

    qc_process = appsettings.get("QCProcess") or {}
    assert "ProcessTypes" in qc_process
    assert qc_process["ProcessTypes"]["production"]["PdfSuffix"] == "prod"
    assert qc_process["ProcessTypes"]["check"]["PdfSuffix"] == "chk"
    assert qc_process["ProcessTypes"]["review"]["PdfSuffix"] == "rev"

    wf = appsettings["qcWorkflow"]
    assert wf["states"]["production"] == "In Development"
    assert wf["states"]["qcInitiated"] == "Initiate Origination"
    assert wf["states"]["readyForQc"] == "Originated"
    assert wf["attributeMap"]["processType"] == "QC_Process_Type"

    dedupe = appsettings["notifications"]["dedupe"]["keyFields"]
    assert "qcProcessType" in dedupe
