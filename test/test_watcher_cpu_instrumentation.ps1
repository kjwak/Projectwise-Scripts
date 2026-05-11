# Integration test: watcher emits CPU/load timing counters and reuses local cache data safely.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Invoke-WatcherAndGetDone($RepoRoot, $AppSettingsPath) {
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$RepoRoot\Watch-QCTrigger.ps1" -AppSettingsPath $AppSettingsPath -DryRun 2>&1
    $lines = @($out | ForEach-Object { $_ -as [string] } | Where-Object { $_ })
    foreach ($line in $lines) {
        if ($line -notmatch '\"code\"\s*:\s*\"WATCH_DONE\"') { continue }
        $evt = $line | ConvertFrom-Json
        if ($evt.code -eq 'WATCH_DONE') { return $evt }
    }

    Write-Host 'Captured watcher output:' -ForegroundColor Yellow
    $lines | Select-Object -First 50 | ForEach-Object { Write-Host $_ }
    throw 'ASSERT FAILED: watcher should emit WATCH_DONE'
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$tempRoot = Join-Path $env:TEMP ("qc-watch-cpu-test-" + ([guid]::NewGuid().ToString('N')))
$watchDir = Join-Path $tempRoot 'watch'
$queueDir = Join-Path $tempRoot 'queue'
New-Item -ItemType Directory -Path $watchDir -Force | Out-Null
New-Item -ItemType Directory -Path $queueDir -Force | Out-Null

try {
    Set-Content -LiteralPath (Join-Path $watchDir 'A101.pdf') -Value 'dummy' -Encoding utf8

    $appSettingsPath = Join-Path $tempRoot 'appsettings.json'
    $config = @{
        dryRun = $true
        watchFolders = @($watchDir)
        queue = @{ rootDir = $queueDir }
        watcher = @{ idleSleepMs = 1 }
        triggers = @{
            rules = @(
                @{
                    id = 'qc-prepend-local'
                    enabled = $true
                    priority = 100
                    jobType = 'QC_PREPEND'
                    triggerType = 'fs'
                    when = @{ extensions = @('.pdf'); fileNameRegexAny = @('(?i)\.pdf$') }
                    requireAll = @('extensions','fileNameRegexAny')
                    exclude = @{ pathRegexAny = @(); fileNameRegexAny = @() }
                    grouping = @{ enabled = $false; groupBy = 'file' }
                }
            )
        }
        filters = @{ whitelist = @{ enabled = $false; paths = @() }; blacklist = @{ paths = @(); patterns = @() } }
    }
    ($config | ConvertTo-Json -Depth 50) | Set-Content -LiteralPath $appSettingsPath -Encoding utf8

    $done = Invoke-WatcherAndGetDone -RepoRoot $repoRoot -AppSettingsPath $appSettingsPath
    Assert-True ($done.data.elapsedMs -ge 0) 'WATCH_DONE should include elapsedMs'
    Assert-True ($null -ne $done.data.phaseMs) 'WATCH_DONE should include phaseMs'
    Assert-True ($null -ne $done.data.phaseCounts) 'WATCH_DONE should include phaseCounts'
    Assert-True ($done.data.phaseCounts.localFilesDiscovered -eq 1) 'phaseCounts should include discovered local file count'
    Assert-True ($done.data.phaseCounts.hashesCalculated -eq 1) 'first pass should calculate one file hash'
    Assert-True ($done.data.phaseCounts.hashCacheMisses -eq 1) 'first pass should miss hash cache'
    Assert-True ($done.data.phaseCounts.localCacheMisses -eq 1) 'first pass should miss local signature cache'
    Assert-True ($done.data.phaseCounts.dedupeChecks -eq 1) 'phaseCounts should count dedupe checks'
    Assert-True ($done.data.phaseMs.candidateDiscoveryMs -ge 0) 'phaseMs should include candidate discovery timing'
    Assert-True ($done.data.phaseMs.hashCalculationMs -ge 0) 'phaseMs should include hash timing'

    # Dry-run must preserve evaluation/logging semantics, but unchanged file hash should be memoized.
    $done2 = Invoke-WatcherAndGetDone -RepoRoot $repoRoot -AppSettingsPath $appSettingsPath
    Assert-True ($done2.data.phaseCounts.localCacheHits -eq 1) 'second pass should hit local signature cache'
    Assert-True ($done2.data.phaseCounts.localCacheSkips -eq 0) 'dry-run should not skip accepted unchanged files'
    Assert-True ($done2.data.phaseCounts.hashCacheHits -eq 1) 'second pass should reuse cached hash'
    Assert-True ($done2.data.phaseCounts.hashesCalculated -eq 0) 'second pass should avoid rehashing unchanged file'
    Assert-True ($done2.data.phaseCounts.dedupeChecks -eq 1) 'dry-run should still perform dedupe check for wouldDedupe logging'

    Write-Host 'test_watcher_cpu_instrumentation: passed' -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
