$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_inflight_' + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path (Join-Path $tmp 'running') -Force | Out-Null
$config = @{ queue = @{ rootDir = $tmp } }

$folder = 'Documents\AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Sheets'
$job = @{
    id = 'qc_statussetgen_test'
    type = 'STATUS_SET_GEN'
    sourceFolder = $folder.ToLowerInvariant()
    status = 'running'
}
$jobPath = Join-Path $tmp 'running\qc_statussetgen_test.json'
$job | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jobPath -Encoding utf8

$hit = Test-QCStatusSetJobInFlight -Config $config -SourceFolder $folder
if (-not $hit.IsSuccess) { throw "expected success: $($hit.Message)" }
if (-not [bool]$hit.Data.inFlight) { throw 'expected inFlight=true' }
if ([string]$hit.Data.queueState -ne 'running') { throw "expected running, got $($hit.Data.queueState)" }

$miss = Test-QCStatusSetJobInFlight -Config $config -SourceFolder 'Documents\Other\CADD\Sheets'
if (-not $miss.IsSuccess -or [bool]$miss.Data.inFlight) { throw 'expected no match for other folder' }

Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'PASS test_statusset_job_in_flight.ps1'
