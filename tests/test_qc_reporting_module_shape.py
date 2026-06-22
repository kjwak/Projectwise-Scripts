from pathlib import Path

from module_impl import module_impl_path

REPO_ROOT = Path(__file__).resolve().parents[1]
REPORTING = module_impl_path("QC.Reporting.psm1")
PROCESSORS = module_impl_path("QC.Processors.psm1")


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
        "qcClosedCount",
        "qcErrorCount",
        "staleQcCount",
        "inProductionCount",
        "readyForQcCount",
        "redlinesReceivedCount",
        "correctionsReceivedCount",
        "qcFinalizingCount",
        "reviewInProgressCount",
        "redlinesIssuedCount",
        "correctionsInProgressCount",
        "verificationInProgressCount",
        "qcCompleteCount",
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
