from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = (REPO / 'modules' / 'Core.Database.psm1').read_text(encoding='utf-8')
WATCHER = (REPO / 'scripts' / 'Watch-QCTrigger.ps1').read_text(encoding='utf-8')


def test_validation_helper_present() -> None:
    assert 'function Test-QCPollRunTelemetryInsertShape' in DB
    assert 'POLL_RUN_TELEMETRY_COLUMN_VALUE_MISMATCH' in DB
    assert 'POLL_RUN_TELEMETRY_PARAM_REF_MISSING' in DB


def test_write_function_returns_error_result_and_logs() -> None:
    assert 'POLL_RUN_TELEMETRY_WRITE_FAILED' in DB
    assert 'return New-QCErrorResult -Code' in DB


def test_watcher_uses_counterpath_and_fail_on_write_error() -> None:
    assert 'Get-AuditPollCycleCounter -CounterPath $counterPath' in WATCHER
    assert 'WATCH_PASS_COUNTER_READ_FAILED' in WATCHER
    assert 'telemetry.failOnWriteError' in WATCHER
    assert 'if ($telemetryFailOnWriteError) { throw' in WATCHER


def test_poll_run_validation_not_used_for_job_or_sheet_telemetry() -> None:
    job_fn = DB.split('function Write-QCJobTelemetry', 1)[1].split('function Test-QCPollRunTelemetryInsertShape', 1)[0]
    sheet_fn = DB.split('function Write-QCSheetIndex', 1)[1].split('function Update-QCSheetIndexPwStateName', 1)[0]
    assert 'Test-QCPollRunTelemetryInsertShape' not in job_fn
    assert 'Test-QCPollRunTelemetryInsertShape' not in sheet_fn
    poll_fn = DB.split('function Write-QCPollRunTelemetry', 1)[1].split('function Write-QCNotificationTelemetry', 1)[0]
    assert 'Test-QCPollRunTelemetryInsertShape' in poll_fn
