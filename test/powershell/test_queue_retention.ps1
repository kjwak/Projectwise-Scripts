# Tests recon-slot cleanup of aged succeeded jobs and queue\_logs files.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Import-Module "$PSScriptRoot/../../modules/Core/Core.Results.psm1" -Force
Import-Module "$PSScriptRoot/../../modules/Queue/QC.Queue.Json.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

function New-AgedFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][datetime]$LastWriteUtc,
        [string]$Content = 'x'
    )
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Set-Content -LiteralPath $Path -Value $Content -Encoding UTF8
    $item = Get-Item -LiteralPath $Path
    $item.LastWriteTimeUtc = $LastWriteUtc
}

$root = Join-Path $env:TEMP ("qc-queue-ret-" + ([guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $root -Force | Out-Null

try {
    foreach ($sub in @('pending', 'running', 'succeeded', 'failed', 'locks', '_logs', '_watcher')) {
        New-Item -ItemType Directory -Path (Join-Path $root $sub) -Force | Out-Null
    }

    $now = [DateTime]::UtcNow
    $old = $now.AddHours(-25)
    $recent = $now.AddHours(-1)

    New-AgedFile -Path (Join-Path $root 'succeeded\job-old.json') -LastWriteUtc $old -Content '{"id":"job-old","dedupeKey":"dq_old"}'
    New-AgedFile -Path (Join-Path $root 'succeeded\job-new.json') -LastWriteUtc $recent -Content '{"id":"job-new","dedupeKey":"dq_new"}'
    New-AgedFile -Path (Join-Path $root 'failed\job-failed.json') -LastWriteUtc $old -Content '{"id":"job-failed"}'
    New-AgedFile -Path (Join-Path $root '_logs\Watch-QCTrigger_2026-01-01_06.jsonl') -LastWriteUtc $old
    New-AgedFile -Path (Join-Path $root '_logs\worker-current.stdout.log') -LastWriteUtc $recent
    New-AgedFile -Path (Join-Path $root '_logs\nested\old-stderr.log') -LastWriteUtc $old

    $idxPath = Join-Path $root '_watcher\dedupe-index.json'
    @{
        version = 1
        entries = @{
            dq_old = @{ jobId = 'job-old'; state = 'succeeded'; updatedAtUtc = $old.ToString('o') }
            dq_new = @{ jobId = 'job-new'; state = 'succeeded'; updatedAtUtc = $recent.ToString('o') }
        }
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $idxPath -Encoding UTF8

    $config = @{
        queue = @{
            rootDir = $root
            retention = @{ enabled = $true; succeededHours = 24; logsHours = 24 }
        }
    }

    $defaults = Get-QCQueueRetentionSettings -Config @{ queue = @{ rootDir = $root } }
    Assert-True $defaults.enabled 'Missing retention section should default enabled'
    Assert-Eq $defaults.succeededHours 24 'Default succeededHours should be 24'
    Assert-Eq $defaults.logsHours 24 'Default logsHours should be 24'

    $dry = Invoke-QCQueueRetention -Config $config -DryRun
    Assert-True $dry.IsSuccess 'Dry-run should succeed'
    Assert-Eq $dry.Data.succeeded.deleted 1 'Dry-run should count one old succeeded job'
    Assert-True (Test-Path -LiteralPath (Join-Path $root 'succeeded\job-old.json')) 'Dry-run must not delete succeeded files'
    Assert-True (Test-Path -LiteralPath (Join-Path $root '_logs\Watch-QCTrigger_2026-01-01_06.jsonl')) 'Dry-run must not delete log files'

    $live = Invoke-QCQueueRetention -Config $config
    Assert-True $live.IsSuccess 'Live retention should succeed'
    Assert-Eq $live.Data.succeeded.deleted 1 'Should delete one old succeeded job'
    Assert-Eq $live.Data.logs.deleted 2 'Should delete two old log files (root + nested)'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $root 'succeeded\job-old.json'))) 'Old succeeded job should be gone'
    Assert-True (Test-Path -LiteralPath (Join-Path $root 'succeeded\job-new.json')) 'Recent succeeded job should remain'
    Assert-True (Test-Path -LiteralPath (Join-Path $root 'failed\job-failed.json')) 'Failed jobs must not be pruned'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $root '_logs\Watch-QCTrigger_2026-01-01_06.jsonl'))) 'Old jsonl should be gone'
    Assert-True (Test-Path -LiteralPath (Join-Path $root '_logs\worker-current.stdout.log')) 'Recent log should remain'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $root '_logs\nested\old-stderr.log'))) 'Nested old log should be gone'
    Assert-Eq $live.Data.dedupeIndexPruned 1 'Dedupe index should drop the deleted succeeded job'

    $dupOld = Test-QCDuplicateJob -DedupeKey 'dq_old' -Config $config
    Assert-True ($dupOld.IsSuccess -and -not $dupOld.Data.isDuplicate) 'Deleted succeeded job should not stay a duplicate'
    $dupNew = Test-QCDuplicateJob -DedupeKey 'dq_new' -Config $config
    Assert-True ($dupNew.IsSuccess -and $dupNew.Data.isDuplicate) 'Recent succeeded job should still be a duplicate'

    $skip = Invoke-QCQueueRetention -Config @{
        queue = @{ rootDir = $root; retention = @{ enabled = $false } }
    }
    Assert-Eq $skip.Code 'QUEUE_RETENTION_SKIPPED' 'Disabled retention should skip'

    Write-Host 'test_queue_retention.ps1: all assertions passed.' -ForegroundColor Green
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
