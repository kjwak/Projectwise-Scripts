from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
REPORTING = REPO_ROOT / "modules" / "QC.Reporting.psm1"
PROCESSORS = REPO_ROOT / "modules" / "QC.Processors.psm1"


def test_reporting_module_exports_expected_functions_and_metrics():
    text = REPORTING.read_text(encoding="utf-8")
    for name in [
        "Get-QCReportingSettings",
        "Get-QCReportingDocuments",
        "ConvertTo-QCReportingDocument",
        "New-QCReportingSnapshot",
        "Write-QCReportingSnapshot",
        "Invoke-QCReportingScan",
        "New-QCReportingScanJob",
    ]:
        assert name in text
    for metric in [
        "qcActiveCount",
        "qcOpenCount",
        "qcPendingBackcheckCount",
        "qcClosedCount",
        "qcErrorCount",
        "staleQcCount",
        "inProductionCount",
        "qcReceivedCount",
        "redlinesIssuedCount",
        "correctionsInProgressCount",
        "correctionsCompleteCount",
        "backcheckInProgressCount",
        "verifiedClosedCount",
        "errorNeedsAttentionCount",
        "staleOpenQcCount",
        "avgQcCycleDays",
    ]:
        assert metric in text


def test_reporting_processor_dispatch_is_registered():
    text = PROCESSORS.read_text(encoding="utf-8")
    assert "QC_REPORTING_SCAN" in text
    assert "Invoke-QCReportingScanProcessor" in text
    assert "QC.Reporting.psm1" in text


def test_reporting_module_is_read_only_for_projectwise():
    text = REPORTING.read_text(encoding="utf-8")
    forbidden = [
        "Set-PW",
        "Update-PW",
        "New-PW",
        "Remove-PW",
        "Rename-PW",
    ]
    for token in forbidden:
        assert token not in text
