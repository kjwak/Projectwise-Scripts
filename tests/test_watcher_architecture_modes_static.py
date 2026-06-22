from pathlib import Path

from module_impl import read_module_source

REPO = Path(__file__).resolve().parents[1]
WATCHER = (REPO / 'scripts' / 'Watch-QCTrigger.ps1').read_text(encoding='utf-8')
DB = read_module_source('Core.Database.psm1')
ORCH = read_module_source('QC.WatcherOrchestration.psm1')


def test_orchestration_module_defines_modes_and_plan() -> None:
    assert "function Get-QCWatcherMode" in ORCH
    assert "function Get-QCReconciliationPlan" in ORCH
    assert "function Get-QCWatcherContinuousSettings" in ORCH
    assert "audit_only" in ORCH and "reconciliation" in ORCH and "recovery" in ORCH and "hybrid" in ORCH


def test_watcher_supports_continuous_pw_session() -> None:
    assert "Get-QCWatcherContinuousSettings" in WATCHER
    assert "-Continuous" in WATCHER
    assert "WATCH_TICK_START" in WATCHER
    assert "WATCH_PW_DISCONNECT" in WATCHER
    assert "} while ($watcherContinuous)" in WATCHER


def test_watcher_imports_orchestration_module() -> None:
    assert "Import-Module (Join-Path $repoRoot 'modules\\QC.WatcherOrchestration.psm1')" in WATCHER
    assert "Get-QCWatcherMode -Config $config" in WATCHER
    assert "Get-QCReconciliationPlan -Config $config" in WATCHER


def test_poll_runs_schema_supports_reconciliation_lifecycle_fields() -> None:
    for token in (
        'reconciliation_trigger_source NVARCHAR(80)',
        'downtime_seconds INT',
        'audit_gap_detected BIT NOT NULL DEFAULT 0',
        'watcher_phase NVARCHAR(80)',
        'throttle_wait_seconds DECIMAL(18,3)',
        'queue_depth_snapshot INT',
    ):
        assert token in DB
