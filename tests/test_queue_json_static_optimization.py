from pathlib import Path

MODULE = Path(__file__).resolve().parents[1] / "modules" / "QC.Queue.Json.psm1"
SOURCE = MODULE.read_text(encoding="utf-8")


def test_queue_json_array_conversions_use_linear_lists() -> None:
    assert "Queue jobs can contain large paired-sheet" in SOURCE
    assert "when reading jobs back from JSON" in SOURCE
    assert "[System.Collections.Generic.List[object]]::new()" in SOURCE
    assert "$out += _QCQJ-DeepToJsonSafeObject" not in SOURCE
    assert "$out += (_ToHashtable $i)" not in SOURCE


def test_get_next_qc_job_uses_single_pass_preference_scheduler() -> None:
    assert "Read each eligible job at most once" in SOURCE
    assert "$preferRank = @{}" in SOURCE
    assert "$bestPreferredRank = [int]::MaxValue" in SOURCE
    assert "if ($rank -eq 0)" in SOURCE
    assert "foreach ($pt in $preferTypes)" not in SOURCE


def test_watcher_active_flag_self_heals_dead_owner_pid() -> None:
    assert "Self-heal stale watcher-active flags" in SOURCE
    assert "Get-Process -Id $ownerPid" in SOURCE
    assert "Remove-Item -LiteralPath $flag" in SOURCE
