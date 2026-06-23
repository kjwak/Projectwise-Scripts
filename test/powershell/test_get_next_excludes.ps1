# Unit test: Get-NextQCJob -ExcludeJobIds skips supplied ids and returns next eligible.
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

$tempRoot = Join-Path $env:TEMP ("qc-excl-test-" + ([guid]::NewGuid().ToString('N')))
$queueDir = Join-Path $tempRoot 'queue'
New-Item -ItemType Directory -Path $queueDir -Force | Out-Null

try {
    $cfg = @{ queue = @{ rootDir = $queueDir } }

    foreach ($id in @('aaa','bbb','ccc')) {
        $job = @{
            id = $id
            type = 'STATUS_SET_GEN'
            sourcePath = "c:\\work\\$id"
            sourceFolder = "c:\\work"
            sourceName = $id
            dedupeKey = "dq_$id"
            status = 'pending'
            attempts = 0
            metadata = @{}
        }
        $r = Add-QCQueueJob -Job $job -Config $cfg
        Assert-True $r.IsSuccess "enqueue $id"
        # Brief sleep so LastWriteTimeUtc ordering is deterministic across enqueues.
        Start-Sleep -Milliseconds 25
    }

    $first = Get-NextQCJob -Config $cfg
    Assert-True $first.IsSuccess 'first selection'
    Assert-True ($null -ne $first.Data.job) 'first job present'
    $firstId = [string]$first.Data.jobId
    Assert-True ($firstId -in @('aaa','bbb','ccc')) "first id valid: $firstId"

    $second = Get-NextQCJob -Config $cfg -ExcludeJobIds @($firstId)
    Assert-True $second.IsSuccess 'second selection'
    Assert-True ($null -ne $second.Data.job) 'second job present'
    $secondId = [string]$second.Data.jobId
    Assert-True ($secondId -ne $firstId) 'exclude must skip the first id'

    $third = Get-NextQCJob -Config $cfg -ExcludeJobIds @($firstId, $secondId)
    Assert-True $third.IsSuccess 'third selection'
    Assert-True ($null -ne $third.Data.job) 'third job present'
    $thirdId = [string]$third.Data.jobId
    Assert-True ($thirdId -ne $firstId -and $thirdId -ne $secondId) 'exclude must skip both prior ids'

    $empty = Get-NextQCJob -Config $cfg -ExcludeJobIds @($firstId, $secondId, $thirdId)
    Assert-True $empty.IsSuccess 'empty selection'
    Assert-Eq $empty.Data.job $null 'when all ids excluded, no job returned'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'test_get_next_excludes: passed' -ForegroundColor Green
