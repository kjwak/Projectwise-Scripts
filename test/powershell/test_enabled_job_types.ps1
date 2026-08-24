# Unit test: workers.enabledJobTypes / Get-NextQCJob -IncludeJobTypes and UNC claim guard.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.Queue.Json.psm1" -Force

$tempRoot = Join-Path $env:TEMP ("qc-enabled-types-" + ([guid]::NewGuid().ToString('N')))
$queueDir = Join-Path $tempRoot 'queue'
New-Item -ItemType Directory -Path $queueDir -Force | Out-Null

try {
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

    $cfg = @{ queue = @{ rootDir = $queueDir } }
    Add-QCQueueJob -Job (_Job 'ssg-1' 'STATUS_SET_GEN') -Config $cfg | Out-Null
    Start-Sleep -Milliseconds 25
    Add-QCQueueJob -Job (_Job 'qcp-1' 'QC_PREPEND') -Config $cfg | Out-Null

    $all = Get-NextQCJob -Config $cfg
    Assert-True $all.IsSuccess 'unfiltered selection succeeded'
    Assert-True ($null -ne $all.Data.job) 'unfiltered returned a job'
    Assert-Eq ([string]$all.Data.jobId) 'ssg-1' 'unfiltered prefers oldest (STATUS_SET_GEN)'

    $onlyPrepend = Get-NextQCJob -Config $cfg -IncludeJobTypes @('QC_PREPEND')
    Assert-True $onlyPrepend.IsSuccess 'include QC_PREPEND succeeded'
    Assert-Eq ([string]$onlyPrepend.Data.jobId) 'qcp-1' 'include must skip STATUS_SET_GEN'

    $none = Get-NextQCJob -Config $cfg -IncludeJobTypes @('QC_RENDITION')
    Assert-True $none.IsSuccess 'include unmatched type succeeded'
    Assert-Eq $none.Data.job $null 'include unmatched type returns no job'

    $combo = Get-NextQCJob -Config $cfg -IncludeJobTypes @('QC_PREPEND', 'STATUS_SET_GEN') -ExcludeJobTypes @('STATUS_SET_GEN')
    Assert-Eq ([string]$combo.Data.jobId) 'qcp-1' 'include+exclude must leave QC_PREPEND'

    $emptyInclude = Get-NextQCJob -Config $cfg -IncludeJobTypes @()
    Assert-Eq ([string]$emptyInclude.Data.jobId) 'ssg-1' 'empty include list means all types'

    $unrestricted = @(Get-QCEnabledJobTypes -Config $cfg)
    Assert-Eq $unrestricted.Count 0 'missing enabledJobTypes means unrestricted'

    $cfgFilter = @{
        queue = @{ rootDir = $queueDir }
        workers = @{ enabledJobTypes = @('QC_PREPEND') }
    }
    $fromCfg = @(Get-QCEnabledJobTypes -Config $cfgFilter)
    Assert-Eq $fromCfg.Count 1 'one enabled type'
    Assert-Eq $fromCfg[0] 'QC_PREPEND' 'enabled type is QC_PREPEND'

    $rh = Get-QCRemoteWorkerHostSettings -Config $cfgFilter
    $rhTypes = @($rh.enabledJobTypes)
    Assert-Eq $rhTypes.Count 1 'remote host settings expose enabledJobTypes'
    Assert-Eq $rhTypes[0] 'QC_PREPEND' 'remote host enabled type is QC_PREPEND'
    Assert-Eq ([bool]$rh.allowUncQueue) $false 'allowUncQueue defaults false'

    Assert-True (Test-QCQueueRootIsUnc -Path '\\192.168.22.90\QC_Queue') 'UNC share is UNC'
    Assert-True (-not (Test-QCQueueRootIsUnc -Path 'C:\QC_E2E_RealRun\queue')) 'local path is not UNC'
    Assert-True (Test-QCUncQueueClaimAllowed -Config $cfg) 'local queue claims allowed'
    $uncCfg = @{
        queue = @{ rootDir = '\\192.168.22.90\QC_Queue' }
        workers = @{ remoteHost = @{ allowUncQueue = $false } }
    }
    Assert-True (-not (Test-QCUncQueueClaimAllowed -Config $uncCfg)) 'UNC refused without opt-in'
    Assert-True (Test-QCUncQueueClaimAllowed -Config $uncCfg -AllowUncQueue) 'UNC allowed with switch'
    $uncCfg.workers.remoteHost.allowUncQueue = $true
    Assert-True (Test-QCUncQueueClaimAllowed -Config $uncCfg) 'UNC allowed via config flag'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$hostScript = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Start-QCRemoteWorkerHost.ps1') -Raw
$logView = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\QC.RemoteWorkerHostLogView.ps1') -Raw
Assert-True ($hostScript -match 'Drain-QCRemoteWorkerHostJsonLogs') 'remote host tails processor JSONL'
Assert-True ($logView -match 'WORKER_SELECTED') 'remote host prints WORKER_SELECTED'
Assert-True ($hostScript -match 'Tailing processor JSON logs') 'remote host announces console job activity'

Write-Host 'test_enabled_job_types: passed' -ForegroundColor Green
