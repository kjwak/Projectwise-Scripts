# Minimal integration test for Watch-QCTrigger.ps1 dry-run logging semantics.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Core/Core.Paths.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.Triggers.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.JobFactory.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.Queue.Json.psm1" -Force

$tempRoot = Join-Path $env:TEMP ("qc-watch-dryrun-test-" + ([guid]::NewGuid().ToString('N')))
$watchDir = Join-Path $tempRoot "watch"
$queueDir = Join-Path $tempRoot "queue"
New-Item -ItemType Directory -Path $watchDir -Force | Out-Null
New-Item -ItemType Directory -Path $queueDir -Force | Out-Null

try {
    # Create a dummy file.
    $filePath = Join-Path $watchDir "A101.pdf"
    Set-Content -LiteralPath $filePath -Value "dummy" -Encoding utf8

    $appSettingsPath = Join-Path $tempRoot "appsettings.json"
    $config = @{
        dryRun = $true
        watchFolders = @($watchDir)
        queue = @{ rootDir = $queueDir }
        triggers = @{
            rules = @(
                @{
                    id = 'qc-prepend-local'
                    enabled = $true
                    priority = 100
                    jobType = 'QC_PREPEND'
                    triggerType = 'fs'
                    when = @{ extensions = @('.pdf'); descriptionContainsAny = @(); pathRegexAny = @(); fileNameRegexAny = @('(?i)\.pdf$') }
                    requireAll = @('extensions','fileNameRegexAny')
                    exclude = @{ pathRegexAny = @(); fileNameRegexAny = @() }
                    grouping = @{ enabled = $false; groupBy = 'file' }
                }
            )
        }
        filters = @{ whitelist = @{ enabled = $false; paths = @() }; blacklist = @{ paths = @(); patterns = @() } }
    }
    ($config | ConvertTo-Json -Depth 50) | Set-Content -LiteralPath $appSettingsPath -Encoding utf8

    # Seed the queue with a duplicate job matching the same dedupeKey.
    $norm = Normalize-QCPath -Path $filePath
    Assert-True $norm.IsSuccess 'Normalize-QCPath should succeed'
    $sha = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    $candidate = @{
        path = [string]$norm.Data.path
        fileName = 'A101.pdf'
        description = ''
        detectedAtUtc = '2026-01-01T00:00:00.0000000Z'
        sourceFolder = (Normalize-QCPath -Path $watchDir).Data.path
        file = @{ fullName = $filePath; sha256 = $sha }
    }
    $rule = $config.triggers.rules[0]
    $m = Resolve-QCTriggerMatch -Candidate $candidate -Config $config
    Assert-True ($m.IsSuccess -and $m.Data.matched) 'Test fixture trigger rule should match candidate'
    $jobRes = New-QCJobObject -Candidate $candidate -Rule $rule -Config $config
    Assert-True $jobRes.IsSuccess 'New-QCJobObject should succeed'
    $enq = Add-QCQueueJob -Job $jobRes.Data.job -Config $config
    Assert-True $enq.IsSuccess 'Add-QCQueueJob should succeed'

    # Run watcher in dry-run. It should report wouldDedupe=true and reason=duplicate.
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\\Watch-QCTrigger.ps1" -AppSettingsPath $appSettingsPath -DryRun 2>&1
    $lines = @($out | ForEach-Object { $_ -as [string] } | Where-Object { $_ })
    $accepted = $null
    foreach ($l in $lines) {
        if ($l -notmatch '\"code\"\s*:\s*\"WATCH_ACCEPTED\"') { continue }
        $o = $l | ConvertFrom-Json
        if ($o.code -eq 'WATCH_ACCEPTED') { $accepted = $o; break }
    }
    if ($null -eq $accepted) {
        Write-Host "Captured watcher output (first 50 lines):" -ForegroundColor Yellow
        $i = 0
        foreach ($l in $lines) {
            $i++
            Write-Host $l
            if ($i -ge 50) { break }
        }
        throw "ASSERT FAILED: Should emit WATCH_ACCEPTED log"
    }
    Assert-Eq $accepted.data.dryRun $true 'dryRun should be true in log'
    Assert-Eq $accepted.data.wouldDedupe $true 'wouldDedupe should be true for seeded duplicate'
    Assert-Eq $accepted.data.wouldEnqueue $false 'wouldEnqueue should be false when duplicate'
    Assert-Eq $accepted.data.enqueueSkippedReason 'duplicate' 'enqueueSkippedReason should be duplicate'

    Write-Host 'All Watch-QCTrigger dry-run logging tests passed.' -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

