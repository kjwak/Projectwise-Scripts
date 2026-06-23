# Unit test: STATUS_SET_GEN is gated while the watcher-active flag is set;
# QC_PREPEND remains immediately eligible.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.Queue.Json.psm1" -Force

$tempRoot = Join-Path $env:TEMP ("qc-gate-test-" + ([guid]::NewGuid().ToString('N')))
$queueDir = Join-Path $tempRoot 'queue'
New-Item -ItemType Directory -Path $queueDir -Force | Out-Null

try {
    $cfg = @{ queue = @{ rootDir = $queueDir } }

    function _Job($Id, $Type) {
        return @{
            id = $Id
            type = $Type
            sourcePath = "c:\\work\\$Id"
            sourceFolder = "c:\\work"
            sourceName = $Id
            dedupeKey = "dq_$Id"
            status = 'pending'
            attempts = 0
            metadata = @{}
        }
    }

    Add-QCQueueJob -Job (_Job 'ssg-1' 'STATUS_SET_GEN') -Config $cfg | Out-Null
    Start-Sleep -Milliseconds 25
    Add-QCQueueJob -Job (_Job 'qcp-1' 'QC_PREPEND') -Config $cfg | Out-Null

    # Baseline: no flag, no exclusion -> first by enqueue time = ssg-1 (unless preferTypes wins).
    Assert-True (-not (Test-QCWatcherActive -Config $cfg)) 'flag should not exist initially'
    $r0 = Get-NextQCJob -Config $cfg
    Assert-True $r0.IsSuccess 'r0 selection succeeded'
    Assert-True ($null -ne $r0.Data.job) 'r0 returned a job'

    # With flag set -> STATUS_SET_GEN excluded -> must select qcp-1.
    Set-QCWatcherActive -Config $cfg | Out-Null
    Assert-True (Test-QCWatcherActive -Config $cfg) 'flag should exist after Set-QCWatcherActive'
    $rGated = Get-NextQCJob -Config $cfg -ExcludeJobTypes @('STATUS_SET_GEN')
    Assert-True $rGated.IsSuccess 'gated selection succeeded'
    Assert-True ($null -ne $rGated.Data.job) 'gated selection returned a job'
    Assert-Eq ([string]$rGated.Data.jobId) 'qcp-1' 'gated selection must return QC_PREPEND'

    # If only STATUS_SET_GEN exists and flag is set, no job is returned (gated).
    $cfg2Root = Join-Path $tempRoot 'queue2'
    New-Item -ItemType Directory -Path $cfg2Root -Force | Out-Null
    $cfg2 = @{ queue = @{ rootDir = $cfg2Root } }
    Add-QCQueueJob -Job (_Job 'ssg-only' 'STATUS_SET_GEN') -Config $cfg2 | Out-Null
    Set-QCWatcherActive -Config $cfg2 | Out-Null
    $rEmpty = Get-NextQCJob -Config $cfg2 -ExcludeJobTypes @('STATUS_SET_GEN')
    Assert-True $rEmpty.IsSuccess 'empty (gated) selection succeeded'
    Assert-Eq $rEmpty.Data.job $null 'gated must return no job when only STATUS_SET_GEN pending'

    # After clearing the flag, STATUS_SET_GEN becomes eligible again.
    Clear-QCWatcherActive -Config $cfg2 | Out-Null
    Assert-True (-not (Test-QCWatcherActive -Config $cfg2)) 'flag should be cleared'
    $rAfter = Get-NextQCJob -Config $cfg2
    Assert-True $rAfter.IsSuccess 'post-clear selection succeeded'
    Assert-True ($null -ne $rAfter.Data.job) 'post-clear returned a job'
    Assert-Eq ([string]$rAfter.Data.jobId) 'ssg-only' 'post-clear must return STATUS_SET_GEN'

    # Preferred-type list still respects exclusion: with preferJobTypes=[STATUS_SET_GEN,QC_PREPEND]
    # and STATUS_SET_GEN excluded, the preferred-type loop must skip STATUS_SET_GEN entirely.
    $cfg3Root = Join-Path $tempRoot 'queue3'
    New-Item -ItemType Directory -Path $cfg3Root -Force | Out-Null
    $cfg3 = @{
        queue = @{
            rootDir = $cfg3Root
            selection = @{ preferJobTypes = @('STATUS_SET_GEN','QC_PREPEND') }
        }
    }
    Add-QCQueueJob -Job (_Job 'ssg-pref' 'STATUS_SET_GEN') -Config $cfg3 | Out-Null
    Start-Sleep -Milliseconds 25
    Add-QCQueueJob -Job (_Job 'qcp-pref' 'QC_PREPEND') -Config $cfg3 | Out-Null
    $rPref = Get-NextQCJob -Config $cfg3 -ExcludeJobTypes @('STATUS_SET_GEN')
    Assert-True $rPref.IsSuccess 'preferred-with-exclusion succeeded'
    Assert-Eq ([string]$rPref.Data.jobId) 'qcp-pref' 'with STATUS_SET_GEN excluded, preferred must fall through to QC_PREPEND'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'test_watcher_gate: passed' -ForegroundColor Green
