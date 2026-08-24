# Host resource throttle: config, matching, fail-open status, slots, supervisor/worker gates.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}
function Assert-False($Condition, $Message) {
    if ($Condition) { throw "ASSERT FAILED: $Message" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.Queue.Json.psm1" -Force
Import-Module "$repoRoot/modules/Queue/QC.HostThrottle.psm1" -Force

$enabledSettings = @{
    enabled = $true
    sampleSeconds = 10
    cpuPercent = 0
    memoryPercent = 0
    processNamePatterns = @('OpenRoads*', 'ras*', 'powershell*')
    busyRecommendedSlots = 0
}

# --- Config / defaults ---
$missing = Get-QCRemoteWorkerHostSettings -Config @{ queue = @{ rootDir = 'C:\q' } }
Assert-False $missing.throttle.enabled 'throttle block absent → disabled'
Assert-Eq ([int]$missing.throttle.sampleSeconds) 10 'sampleSeconds default 10'
Assert-Eq ([double]$missing.throttle.cpuPercent) 0 'cpuPercent default 0'
Assert-Eq ([double]$missing.throttle.memoryPercent) 0 'memoryPercent default 0'
Assert-Eq ([int]$missing.throttle.busyRecommendedSlots) 1 'busyRecommendedSlots default 1'
Assert-Eq @($missing.throttle.processNamePatterns).Count 0 'patterns default empty'

$disabledCfg = Get-QCRemoteWorkerHostSettings -Config @{
    workers = @{
        maxParallel = 2
        remoteHost = @{
            throttle = @{
                enabled = $false
                cpuPercent = 80
                memoryPercent = 90
                processNamePatterns = @('ras*')
            }
        }
    }
}
Assert-False $disabledCfg.throttle.enabled 'enabled=false stays disabled'
$disabledEval = ConvertTo-QCHostThrottleEvaluation -Settings $disabledCfg.throttle -MaxParallel 2 `
    -CpuPercent 99 -MemoryPercent 99 -MatchedProcesses @(@{ Name = 'ras.exe' })
Assert-False $disabledEval.pauseNewClaims 'disabled config never pauses'
Assert-Eq $disabledEval.reason 'disabled' 'disabled reason'
Assert-Eq ([int]$disabledEval.recommendedSlots) 2 'disabled recommendedSlots = maxParallel'

$zeroCpu = ConvertTo-QCHostThrottleEvaluation -Settings @{
    enabled = $true; cpuPercent = 0; memoryPercent = 0; processNamePatterns = @(); busyRecommendedSlots = 0
} -MaxParallel 1 -CpuPercent 100 -MemoryPercent 100
Assert-False $zeroCpu.pauseNewClaims 'cpuPercent=0 and memoryPercent=0 are ignored'

# --- Process matching ---
Assert-True (Test-QCHostProcessNameMatch -Name 'OpenRoadsDesigner.exe' -Patterns @('OpenRoads*')) 'wildcard + .exe'
Assert-True (Test-QCHostProcessNameMatch -Name 'openroadsdesigner' -Patterns @('OPENROADS*')) 'case-insensitive'
Assert-True (Test-QCHostProcessNameMatch -Name 'RAS.exe' -Patterns @('ras*')) '.exe stripped on both sides'
Assert-False (Test-QCHostProcessNameMatch -Name 'notepad.exe' -Patterns @('OpenRoads*', 'ras*')) 'no match'

$procs = @(
    [pscustomobject]@{ ProcessId = 11; Name = 'OpenRoadsDesigner.exe'; ParentProcessId = 1 }
    [pscustomobject]@{ ProcessId = 12; Name = 'ras.exe'; ParentProcessId = 1 }
    [pscustomobject]@{ ProcessId = 13; Name = 'notepad.exe'; ParentProcessId = 1 }
)
$multi = @(Select-QCHostMatchedProcesses -Processes $procs -Patterns @('OpenRoads*', 'ras*') -ExcludePids @())
Assert-Eq $multi.Count 2 'two matching processes'
Assert-Eq @(Select-QCHostMatchedProcesses -Processes $procs -Patterns @('Civil3D*') -ExcludePids @()).Count 0 'no match returns empty'

$psProcs = @(
    [pscustomobject]@{ ProcessId = 100; Name = 'powershell.exe'; ParentProcessId = 1 }
    [pscustomobject]@{ ProcessId = 101; Name = 'powershell.exe'; ParentProcessId = 100 }
    [pscustomobject]@{ ProcessId = 200; Name = 'powershell.exe'; ParentProcessId = 1 }
    [pscustomobject]@{ ProcessId = 201; Name = 'pwsh.exe'; ParentProcessId = 1 }
)
$qcTree = @(Select-QCHostMatchedProcesses -Processes $psProcs -Patterns @('powershell*') -ExcludePids @(100))
$qcIds = @($qcTree | ForEach-Object { [int]$_.ProcessId })
Assert-False ($qcIds -contains 100) 'QC supervisor PID excluded'
Assert-False ($qcIds -contains 101) 'QC worker child in process tree excluded'
Assert-True ($qcIds -contains 200) 'non-QC powershell still matches powershell*'
Assert-False ($qcIds -contains 201) 'pwsh does not match powershell* wildcard'
Assert-True ($qcIds.Count -ge 1) 'do not exclude all PowerShell globally'

# --- Resource gates ---
$baseOn = @{
    enabled = $true
    cpuPercent = 70
    memoryPercent = 85
    processNamePatterns = @('ras*')
    busyRecommendedSlots = 0
}
$cpuLow = ConvertTo-QCHostThrottleEvaluation -Settings $baseOn -MaxParallel 2 -CpuPercent 10 -MemoryPercent 10 -MatchedProcesses @()
Assert-False $cpuLow.pauseNewClaims 'CPU below threshold is normal'
Assert-Eq $cpuLow.reason 'normal' 'normal reason'
Assert-Eq ([int]$cpuLow.recommendedSlots) 2 'normal slots = maxParallel'

$cpuHigh = ConvertTo-QCHostThrottleEvaluation -Settings $baseOn -MaxParallel 2 -CpuPercent 70 -MemoryPercent 10 -MatchedProcesses @()
Assert-True $cpuHigh.pauseNewClaims 'CPU at threshold pauses when busyRecommendedSlots is 0'
Assert-Eq $cpuHigh.reason 'cpu_threshold' 'cpu_threshold reason'
Assert-Eq ([int]$cpuHigh.recommendedSlots) 0 'busy slots 0 fully pauses'

$memHigh = ConvertTo-QCHostThrottleEvaluation -Settings $baseOn -MaxParallel 2 -CpuPercent 10 -MemoryPercent 85 -MatchedProcesses @()
Assert-Eq $memHigh.reason 'memory_threshold' 'memory_threshold reason'

$procOnly = ConvertTo-QCHostThrottleEvaluation -Settings $baseOn -MaxParallel 2 -CpuPercent 1 -MemoryPercent 1 -MatchedProcesses @(@{ Name = 'ras.exe' })
Assert-Eq $procOnly.reason 'matched_process' 'process-only busy'
Assert-True $procOnly.pauseNewClaims 'process match pauses'

$multiSig = ConvertTo-QCHostThrottleEvaluation -Settings $baseOn -MaxParallel 2 -CpuPercent 90 -MemoryPercent 90 -MatchedProcesses @(@{ Name = 'ras.exe' })
Assert-Eq $multiSig.reason 'multiple_signals' 'multiple simultaneous signals'

$errEval = ConvertTo-QCHostThrottleEvaluation -Settings $baseOn -MaxParallel 2 -SampleError
Assert-False $errEval.pauseNewClaims 'sample_error does not pause'
Assert-Eq $errEval.reason 'sample_error' 'sample_error reason'
Assert-Eq ([int]$errEval.recommendedSlots) 2 'sample_error keeps maxParallel'

# --- Slots clamp ---
Assert-Eq (Get-QCHostThrottleDesiredSlots -MaxParallel 2 -RecommendedSlots 2) 2 'normal maxParallel'
Assert-Eq (Get-QCHostThrottleDesiredSlots -MaxParallel 2 -RecommendedSlots 0) 0 'busy zero'
Assert-Eq (Get-QCHostThrottleDesiredSlots -MaxParallel 2 -RecommendedSlots -4) 0 'negative clamps to 0'
Assert-Eq (Get-QCHostThrottleDesiredSlots -MaxParallel 2 -RecommendedSlots 99) 2 'above maxParallel clamps'

$busyTwo = ConvertTo-QCHostThrottleEvaluation -Settings @{
    enabled = $true; cpuPercent = 1; memoryPercent = 0; processNamePatterns = @(); busyRecommendedSlots = 1
} -MaxParallel 3 -CpuPercent 50
Assert-Eq ([int]$busyTwo.recommendedSlots) 1 'busyRecommendedSlots honored when busy'
Assert-False $busyTwo.pauseNewClaims 'busyRecommendedSlots 1 keeps claiming instead of turning QC off'

$scaleDefault = ConvertTo-QCHostThrottleEvaluation -Settings @{
    enabled = $true; cpuPercent = 70; memoryPercent = 0; processNamePatterns = @()
} -MaxParallel 5 -CpuPercent 80
Assert-False $scaleDefault.pauseNewClaims 'default busyRecommendedSlots does not pause all claims'
Assert-Eq ([int]$scaleDefault.recommendedSlots) 1 'default busy pool is 1'

$allow = @(Get-QCHostThrottleClaimAllowLabels -CurrentLabels @('RW5', 'RW1', 'RW2') -RecommendedSlots 1)
Assert-Eq $allow.Count 1 'scale-down allow list is N labels'
Assert-Eq $allow[0] 'RW1' 'lowest worker label keeps claiming'
Assert-True (Test-QCHostThrottleWorkerMayClaim -Decision @{ pauseNewClaims = $false; recommendedSlots = 1; claimAllowLabels = @('RW1') } -WorkerLabel 'RW1') 'allowed worker may claim'
Assert-False (Test-QCHostThrottleWorkerMayClaim -Decision @{ pauseNewClaims = $false; recommendedSlots = 1; claimAllowLabels = @('RW1') } -WorkerLabel 'RW5') 'extra worker yields'
Assert-True (Test-QCHostThrottleWorkerMayClaim -Decision @{ pauseNewClaims = $false; recommendedSlots = 1 } -WorkerLabel 'RW5') 'missing allow list fails open'

# --- Status handling ---
$tmp = Join-Path $env:TEMP ('qc-throttle-' + ([guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $statusPath = Join-Path $tmp '_remote_worker.TESTHOST.throttle.json'
    $now = [datetime]::UtcNow

    $written = Write-QCHostThrottleStatus -Path $statusPath -HostName 'TESTHOST' -SampledAtUtc $now -Evaluation @{
        enabled = $true
        pauseNewClaims = $true
        recommendedSlots = 0
        reason = 'matched_process'
        matchedProcesses = @('OpenRoadsDesigner')
        cpuPercent = 42.1
        memoryPercent = 68.7
    }
    $readBack = Read-QCHostThrottleStatus -Path $statusPath
    Assert-True ([bool]$readBack.pauseNewClaims) 'round-trip pauseNewClaims'
    Assert-Eq ([string]$readBack.reason) 'matched_process' 'round-trip reason'
    Assert-Eq ([int]$readBack.recommendedSlots) 0 'round-trip slots'
    Assert-True ([string]$readBack.sampledAtUtc.EndsWith('Z')) 'sampledAtUtc is UTC Z'

    $pausedDecision = Get-QCHostThrottleClaimDecision -Status $readBack -Settings $enabledSettings -MaxParallel 1 -NowUtc $now
    Assert-True $pausedDecision.pauseNewClaims 'valid fresh paused status pauses claims'
    Assert-True $pausedDecision.honored 'fresh paused status is honored'

    $normalStatus = @{
        enabled = $true
        pauseNewClaims = $false
        recommendedSlots = 1
        reason = 'normal'
        sampledAtUtc = $now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }
    $normalDecision = Get-QCHostThrottleClaimDecision -Status $normalStatus -Settings $enabledSettings -MaxParallel 1 -NowUtc $now
    Assert-False $normalDecision.pauseNewClaims 'valid fresh normal status does not pause'

    $missingDecision = Get-QCHostThrottleClaimDecision -Status $null -Settings $enabledSettings -MaxParallel 1 -NowUtc $now
    Assert-False $missingDecision.pauseNewClaims 'missing status fails open'

    $badJson = [string]([char]123) + 'not-json'
    Set-Content -LiteralPath $statusPath -Value $badJson -Encoding ASCII
    $malformed = Read-QCHostThrottleStatus -Path $statusPath
    Assert-True ($null -eq $malformed) 'malformed status reads as null'
    Assert-False ((Get-QCHostThrottleClaimDecision -Status $malformed -Settings $enabledSettings -MaxParallel 1).pauseNewClaims) 'malformed status fails open'

    $stale = @{
        enabled = $true
        pauseNewClaims = $true
        recommendedSlots = 0
        reason = 'matched_process'
        sampledAtUtc = $now.AddSeconds(-90).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }
    Assert-False (Test-QCHostThrottleStatusFresh -Status $stale -SampleSeconds 10 -NowUtc $now) '90s old status is stale at sampleSeconds=10'
    Assert-False ((Get-QCHostThrottleClaimDecision -Status $stale -Settings $enabledSettings -MaxParallel 1 -NowUtc $now).pauseNewClaims) 'stale status fails open'
    Assert-True (Test-QCHostThrottleStatusFresh -Status $normalStatus -SampleSeconds 10 -NowUtc $now) 'fresh status within 30s'

    $sampleErrStatus = @{
        enabled = $true
        pauseNewClaims = $true
        recommendedSlots = 0
        reason = 'sample_error'
        sampledAtUtc = $now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }
    $errDecision = Get-QCHostThrottleClaimDecision -Status $sampleErrStatus -Settings $enabledSettings -MaxParallel 2 -NowUtc $now
    Assert-False $errDecision.pauseNewClaims 'sample_error does not indefinitely pause workers'
    Assert-Eq ([int]$errDecision.recommendedSlots) 2 'sample_error fail-open uses maxParallel'

    $injected = Update-QCHostThrottleStatus -Settings $enabledSettings -StatusPath (Join-Path $tmp 'inj.json') -MaxParallel 1 -HostName 'TESTHOST' -Sample @{
        cpuPercent = 5
        memoryPercent = 5
        processes = @(
            [pscustomobject]@{ ProcessId = 9; Name = 'OpenRoadsDesigner.exe'; ParentProcessId = 1 }
        )
        sampledAtUtc = $now
    }
    Assert-Eq ([string]$injected.evaluation.reason) 'matched_process' 'injected sample evaluates without live CIM'
    Assert-True $injected.evaluation.pauseNewClaims 'injected matching process pauses'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# --- Supervisor spawn plan ---
$pausedSpawn = Get-QCHostThrottleSpawnPlan -CurrentCount 1 -Want 0
Assert-False $pausedSpawn.shouldSpawn 'paused / zero slots → no new child spawning'
Assert-False $pausedSpawn.shouldStopExcess 'decreasing recommended slots does not kill existing workers'
$grow = Get-QCHostThrottleSpawnPlan -CurrentCount 1 -Want 2
Assert-True $grow.shouldSpawn 'increasing slots allows new workers'
Assert-False $grow.shouldStopExcess 'growth still never stops excess'

$hostScript = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Start-QCRemoteWorkerHost.ps1') -Raw -Encoding UTF8
Assert-True ($hostScript -match 'Get-QCHostThrottleSpawnPlan') 'supervisor uses spawn plan'
Assert-True ($hostScript -match 'REMOTE_HOST_THROTTLE') 'supervisor logs throttle state changes'
Assert-True ($hostScript -notmatch 'shouldStopExcess.*Stop-Process') 'throttle path does not Stop-Process excess workers'
Assert-True ($hostScript -match 'throttle=') 'heartbeat includes throttle state'

# --- Worker dequeue gate ---
$workerScript = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Run-QCProcessor.ps1') -Raw -Encoding UTF8
$idxPause = $workerScript.IndexOf('$thDecision = Get-QCHostThrottleClaimDecision')
$idxNext = $workerScript.IndexOf('$next = Get-NextQCJob')
Assert-True ($idxPause -ge 0 -and $idxNext -gt $idxPause) 'pause check is immediately before Get-NextQCJob'
Assert-True ($workerScript -match 'WORKER_THROTTLE_PAUSE') 'paused worker continues without dequeue'
Assert-False ($workerScript -match '_Process-OneJob[\s\S]{0,200}Get-QCHostThrottleClaimDecision') 'in-flight job path does not re-check pause inside processor'

# Simulated worker: pause is only consulted before claim; an owned job always finishes.
$script:claimCalls = 0
$script:finishedJobs = 0
function Invoke-SimulatedWorkerCycle {
    param([bool]$Pause, [bool]$HoldingJob)
    if ($HoldingJob) {
        $script:finishedJobs++
        return $false
    }
    if ($Pause) { return $false }
    $script:claimCalls++
    return $true
}
$holding = $false
$holding = Invoke-SimulatedWorkerCycle -Pause $false -HoldingJob $holding
Assert-True $holding 'resume allows claiming'
$holding = Invoke-SimulatedWorkerCycle -Pause $true -HoldingJob $holding
Assert-False $holding 'in-flight job finishes even when paused'
Assert-Eq $script:finishedJobs 1 'worker already holding a job finishes it'
Assert-Eq $script:claimCalls 1 'paused worker does not call claim after finish'
$holding = Invoke-SimulatedWorkerCycle -Pause $true -HoldingJob $false
Assert-Eq $script:claimCalls 1 'paused worker does not claim again'
$holding = Invoke-SimulatedWorkerCycle -Pause $false -HoldingJob $false
Assert-Eq $script:claimCalls 2 'resume allows claiming again'

Write-Host 'All host throttle tests passed.' -ForegroundColor Green
