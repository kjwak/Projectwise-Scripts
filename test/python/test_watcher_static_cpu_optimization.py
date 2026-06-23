from pathlib import Path

from module_impl import read_module_source

REPO = Path(__file__).resolve().parents[2]
WATCHER = (REPO / "scripts" / "service" / "Watch-QCTrigger.ps1").read_text(encoding="utf-8")
TRIGGERS = read_module_source("QC.Triggers.psm1")
DASHBOARD = (REPO / "scripts" / "service" / "Start-QCPipelineDashboard.ps1").read_text(encoding="utf-8")


def test_watch_done_includes_phase_and_cache_counters() -> None:
    for token in (
        "elapsedMs = [int64]$watchRunSw.ElapsedMilliseconds",
        "phaseMs = $phaseMs",
        "phaseCounts = $phaseCounts",
        "localCacheHits = $localCacheHits",
        "localCacheMisses = $localCacheMisses",
        "localCacheSkips = $localCacheSkips",
        "localCacheEntries = [int]$localWatcherCache.entries.Count",
        "hashCacheHits = $hashCacheHits",
        "hashCacheMisses = $hashCacheMisses",
        "triggerRuleCacheUses = $triggerRuleCacheUses",
    ):
        assert token in WATCHER


def test_local_cache_is_queue_root_acceleration_layer() -> None:
    assert "Join-Path (Join-Path $queueRoot '_watcher') 'local-file-cache.json'" in WATCHER
    assert "_Read-LocalWatcherCache -Path $localCachePath" in WATCHER
    assert "_Write-LocalWatcherCache -Path $localCachePath" in WATCHER
    assert "configHash = $ConfigHash" in WATCHER
    assert "lastWriteTicksUtc = [int64]$FileInfo.LastWriteTimeUtc.Ticks" in WATCHER
    assert "length = [int64]$FileInfo.Length" in WATCHER
    assert "if ($localCacheEntryMatches -and -not $isDryRun)" in WATCHER


def test_hash_and_trigger_rule_memoization_are_wired() -> None:
    assert "[object[]]$OrderedRules = $null" in TRIGGERS
    assert "foreach ($rule in @($OrderedRules))" in TRIGGERS
    assert "Get-OrderedTriggerRules -Config $config" in WATCHER
    assert "-OrderedRules $orderedTriggerRules" in WATCHER
    assert "$candidate.file.sha256 = $cachedSha" in WATCHER
    assert "$candidate.file.sha256 = Get-Sha256FileHex" in WATCHER


def test_dashboard_uses_configured_watcher_backoff() -> None:
    assert "watcherIdleSleepMs" in DASHBOARD
    assert "if ($Cfg.ContainsKey('watcher')" in DASHBOARD
    assert "Start-Sleep -Milliseconds ([int]$wc.watcherIdleSleepMs)" in DASHBOARD
