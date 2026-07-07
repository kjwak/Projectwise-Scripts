from pathlib import Path

DASHBOARD = (Path(__file__).resolve().parents[2] / "scripts" / "service" / "Start-QCPipelineDashboard.ps1").read_text(encoding="utf-8")
README = (Path(__file__).resolve().parents[2] / "scripts" / "README.md").read_text(encoding="utf-8")


def test_dashboard_defaults_to_production_view() -> None:
    assert "[ValidateSet('Production', 'Detailed')]" in DASHBOARD
    assert "[string]$DashboardView = 'Production'" in DASHBOARD
    assert "function _Get-ProductionFrameLines" in DASHBOARD
    assert "if ($DashboardView -eq 'Detailed') { return _Get-DetailedFrameLines -Cfg $Cfg }" in DASHBOARD


def test_production_view_skips_recent_job_fetch() -> None:
    assert "if ($DashboardView -eq 'Detailed')" in DASHBOARD
    assert "Production mode intentionally skips this heavier filesystem read" in DASHBOARD
    assert "Get-QCRecentJobs -Config $Cfg -Limit $recentLimit" in DASHBOARD


def test_production_view_contains_only_critical_sections() -> None:
    for token in (
        "QC Production Dashboard",
        "Queue: pending={0} running={1} failed={2} locks={3} workers={4}/{5}",
        "ProjectWise: ",
        "Workers: {0} active",
        "Warnings/errors: last {0}",
        "Fatal: {0}",
    ):
        assert token in DASHBOARD


def test_detailed_view_is_documented() -> None:
    assert "-DashboardView Production" in README
    assert "-DashboardView Detailed" in README


def test_dashboard_watcher_pass_tracking_ingests_discard_and_ticks() -> None:
    for token in (
        "function _Poll-ChildLogFile",
        "function _Process-ChildJsonLogObject",
        "WATCH_TICK_START",
        "lastDiscardLen",
        "discardTail",
        "-DedupeEvents",
        "hourly JSONL sink is unavailable",
        "$o.data.elapsedMs",
    ):
        assert token in DASHBOARD
    assert "if ($Child.jsonLogDir -and $Child.lastLogHour -and [string]$Child.lastLogHour -ne $hour)" in DASHBOARD
    assert "$prevPath = _Get-ChildJsonLogPath -Child $Child -HourStamp ([string]$Child.lastLogHour)" in DASHBOARD
