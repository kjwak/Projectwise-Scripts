from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
WATCHER = (REPO / 'scripts' / 'Watch-QCTrigger.ps1').read_text(encoding='utf-8')
DB = (REPO / 'modules' / 'Core.Database.psm1').read_text(encoding='utf-8')


def test_poll_runs_schema_contains_extended_runtime_metrics() -> None:
    for token in (
        'run_mode            NVARCHAR(40)',
        'run_status          NVARCHAR(20)',
        'total_duration_seconds DECIMAL(18,3)',
        'audit_query_duration_seconds DECIMAL(18,3)',
        'reconciliation_duration_seconds DECIMAL(18,3)',
        'trigger_eval_duration_seconds DECIMAL(18,3)',
        'dedupe_duration_seconds DECIMAL(18,3)',
        'queue_write_duration_seconds DECIMAL(18,3)',
        'database_write_duration_seconds DECIMAL(18,3)',
        'cleanup_duration_seconds DECIMAL(18,3)',
        'sleep_throttle_duration_seconds DECIMAL(18,3)',
        'candidate_documents_evaluated INT NOT NULL DEFAULT 0',
        'jobs_skipped_dedupe INT NOT NULL DEFAULT 0',
    ):
        assert token in DB


def test_watcher_writes_extended_poll_run_metrics() -> None:
    for token in (
        '-PassNumber $watcherPassNumber',
        '-RunMode $runMode',
        '-TotalDurationSeconds',
        '-AuditQueryDurationSeconds',
        '-ReconciliationDurationSeconds',
        '-TriggerEvalDurationSeconds',
        '-DedupeDurationSeconds',
        '-QueueWriteDurationSeconds',
        '-CleanupDurationSeconds',
        '-SleepThrottleDurationSeconds',
        '-JobsSkippedDedupe $duplicates',
    ):
        assert token in WATCHER


def test_poll_run_insert_values_match_extended_columns() -> None:
    assert '@reconciliationReason, @passNumberSource)' in DB
    assert '@databaseWriteDurationSeconds' in DB
