# QC.HostThrottle.psm1
# Local host resource throttle for remote QC worker hosts.
# Advisory / fail-open: never abort in-flight jobs; never pause indefinitely on bad status.

Set-StrictMode -Version Latest

$script:QCHostThrottlePrevProcesses = $null
$script:QCHostThrottlePrevAt = $null

function ConvertTo-QCHostProcessMatchName {
    <#
    .SYNOPSIS
    Normalize a process name or wildcard for case-insensitive matching (strip .exe).
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $n = $Name.Trim()
    if ($n.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) {
        $n = $n.Substring(0, $n.Length - 4)
    }
    return $n.ToLowerInvariant()
}

function Test-QCHostProcessNameMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [string[]]$Patterns
    )
    $norm = ConvertTo-QCHostProcessMatchName -Name $Name
    if ([string]::IsNullOrWhiteSpace($norm)) { return $false }
    foreach ($pat in @($Patterns)) {
        $p = ConvertTo-QCHostProcessMatchName -Name ([string]$pat)
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if ($norm -like $p) { return $true }
    }
    return $false
}

function Get-QCHostExcludedPidSet {
    <#
    .SYNOPSIS
    Expand supervisor/worker PIDs to their descendant process tree. Does not exclude all powershell.exe.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Processes,
        [int[]]$RootPids
    )
    $byParent = @{}
    foreach ($p in @($Processes)) {
        if ($null -eq $p) { continue }
        $ppid = 0
        $pidVal = 0
        try { $pidVal = [int]$p.ProcessId } catch { $pidVal = 0 }
        try { $ppid = [int]$p.ParentProcessId } catch { $ppid = 0 }
        if ($pidVal -le 0) { continue }
        if (-not $byParent.ContainsKey($ppid)) {
            $byParent[$ppid] = New-Object System.Collections.Generic.List[int]
        }
        [void]$byParent[$ppid].Add($pidVal)
    }
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    $queue = New-Object System.Collections.Queue
    foreach ($id in @($RootPids)) {
        if ($id -gt 0 -and $set.Add($id)) { $queue.Enqueue($id) }
    }
    while ($queue.Count -gt 0) {
        $cur = [int]$queue.Dequeue()
        if ($byParent.ContainsKey($cur)) {
            foreach ($child in @($byParent[$cur])) {
                if ($set.Add([int]$child)) { $queue.Enqueue([int]$child) }
            }
        }
    }
    return $set
}

function Select-QCHostMatchedProcesses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Processes,
        [string[]]$Patterns,
        [int[]]$ExcludePids
    )
    $pats = @($Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($pats.Count -eq 0) { return @() }
    $exclude = Get-QCHostExcludedPidSet -Processes $Processes -RootPids @($ExcludePids)
    $matched = New-Object System.Collections.Generic.List[object]
    foreach ($p in @($Processes)) {
        if ($null -eq $p) { continue }
        $pidVal = 0
        try { $pidVal = [int]$p.ProcessId } catch { $pidVal = 0 }
        if ($pidVal -gt 0 -and $exclude.Contains($pidVal)) { continue }
        $name = ''
        try { $name = [string]$p.Name } catch { $name = '' }
        if (Test-QCHostProcessNameMatch -Name $name -Patterns $pats) {
            $matched.Add([pscustomobject]@{
                ProcessId = $pidVal
                Name      = $name
            }) | Out-Null
        }
    }
    return @($matched.ToArray())
}

function Get-QCHostProcessCpuTicks {
    param($Process)
    if ($null -eq $Process) { return [int64]0 }
    $k = [int64]0
    $u = [int64]0
    try {
        if ($Process.PSObject.Properties['KernelModeTime'] -and $null -ne $Process.KernelModeTime) {
            $k = [int64]$Process.KernelModeTime
        }
    } catch { $k = [int64]0 }
    try {
        if ($Process.PSObject.Properties['UserModeTime'] -and $null -ne $Process.UserModeTime) {
            $u = [int64]$Process.UserModeTime
        }
    } catch { $u = [int64]0 }
    return ($k + $u)
}

function Get-QCHostProcessWorkingSetMb {
    param($Process)
    if ($null -eq $Process) { return 0.0 }
    $bytes = 0.0
    try {
        if ($Process.PSObject.Properties['WorkingSetSize'] -and $null -ne $Process.WorkingSetSize) {
            $bytes = [double]$Process.WorkingSetSize
        }
    } catch { $bytes = 0.0 }
    if ($bytes -le 0) { return 0.0 }
    return [math]::Round($bytes / 1MB, 1)
}

function Get-QCHostMatchedProcessTreePids {
    <#
    .SYNOPSIS
    PIDs of pattern-matched processes plus their descendants (solver children).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Processes,
        [string[]]$Patterns,
        [int[]]$ExcludePids
    )
    $matched = @(Select-QCHostMatchedProcesses -Processes $Processes -Patterns $Patterns -ExcludePids $ExcludePids)
    $roots = @($matched | ForEach-Object { [int]$_.ProcessId } | Where-Object { $_ -gt 0 })
    if ($roots.Count -eq 0) {
        return @{ matched = @($matched); treePids = (New-Object 'System.Collections.Generic.HashSet[int]') }
    }
    return @{
        matched  = @($matched)
        treePids = (Get-QCHostExcludedPidSet -Processes $Processes -RootPids $roots)
    }
}

function Measure-QCHostMatchedProcessUsage {
    <#
    .SYNOPSIS
    CPU % of one logical processor (can exceed 100) and working-set MB for a PID set.
    Requires two snapshots of KernelModeTime/UserModeTime. First sample is not comparable.
    #>
    [CmdletBinding()]
    param(
        $PreviousProcesses,
        $CurrentProcesses,
        $CandidatePids,
        [double]$ElapsedSeconds
    )
    $out = @{
        comparable        = $false
        sumCpuPercent     = 0.0
        maxCpuPercent     = 0.0
        sumWorkingSetMb   = 0.0
        processCount      = 0
    }
    $currById = @{}
    foreach ($p in @($CurrentProcesses)) {
        if ($null -eq $p) { continue }
        $id = 0
        try { $id = [int]$p.ProcessId } catch { $id = 0 }
        if ($id -gt 0) { $currById[$id] = $p }
    }
    $prevById = @{}
    foreach ($p in @($PreviousProcesses)) {
        if ($null -eq $p) { continue }
        $id = 0
        try { $id = [int]$p.ProcessId } catch { $id = 0 }
        if ($id -gt 0) { $prevById[$id] = $p }
    }
    $ids = @()
    if ($CandidatePids -is [System.Collections.IEnumerable] -and -not ($CandidatePids -is [string])) {
        foreach ($id in @($CandidatePids)) {
            try { $ids += [int]$id } catch { }
        }
    }
    $ids = @($ids | Where-Object { $_ -gt 0 } | Select-Object -Unique)
    $out.processCount = $ids.Count
    $sumMb = 0.0
    foreach ($id in $ids) {
        if ($currById.ContainsKey($id)) {
            $sumMb += (Get-QCHostProcessWorkingSetMb -Process $currById[$id])
        }
    }
    $out.sumWorkingSetMb = [math]::Round($sumMb, 1)
    if ($ElapsedSeconds -lt 0.2 -or $prevById.Count -eq 0 -or $ids.Count -eq 0) {
        return $out
    }
    $sumCpu = 0.0
    $maxCpu = 0.0
    $compared = 0
    foreach ($id in $ids) {
        if (-not $currById.ContainsKey($id) -or -not $prevById.ContainsKey($id)) { continue }
        $deltaTicks = (Get-QCHostProcessCpuTicks -Process $currById[$id]) - (Get-QCHostProcessCpuTicks -Process $prevById[$id])
        if ($deltaTicks -lt 0) { continue }
        $cpuPct = ([double]$deltaTicks / 10000000.0) / $ElapsedSeconds * 100.0
        $sumCpu += $cpuPct
        if ($cpuPct -gt $maxCpu) { $maxCpu = $cpuPct }
        $compared++
    }
    if ($compared -le 0) { return $out }
    $out.comparable = $true
    $out.sumCpuPercent = [math]::Round($sumCpu, 1)
    $out.maxCpuPercent = [math]::Round($maxCpu, 1)
    return $out
}

function Get-QCHostThrottleFreshnessSeconds {
    [CmdletBinding()]
    param([int]$SampleSeconds = 10)
    $s = $SampleSeconds
    if ($s -lt 1) { $s = 10 }
    $f = 3 * $s
    if ($f -lt 30) { $f = 30 }
    return [int]$f
}

function Get-QCHostThrottleDesiredSlots {
    [CmdletBinding()]
    param(
        [int]$MaxParallel,
        [int]$RecommendedSlots
    )
    $max = $MaxParallel
    if ($max -lt 0) { $max = 0 }
    $rec = $RecommendedSlots
    if ($rec -lt 0) { $rec = 0 }
    if ($rec -gt $max) { $rec = $max }
    return [int]$rec
}

function Get-QCHostThrottleClaimAllowLabels {
    <#
    .SYNOPSIS
    Lowest N current worker labels may keep claiming when the pool is scaled down.
    #>
    [CmdletBinding()]
    param(
        [string[]]$CurrentLabels,
        [int]$RecommendedSlots
    )
    if ($RecommendedSlots -le 0) { return @() }
    $sorted = @(@($CurrentLabels) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object {
        if ($_ -match '(\d+)$') { [int]$Matches[1] } else { [int]::MaxValue }
    }, { $_ })
    if ($sorted.Count -le $RecommendedSlots) { return @($sorted) }
    return @($sorted | Select-Object -First $RecommendedSlots)
}

function Test-QCHostThrottleWorkerMayClaim {
    <#
    .SYNOPSIS
    False when fully paused (recommendedSlots 0) or this worker is above the scaled-down allow list.
    Missing allow list fails open so workers keep claiming.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Decision,
        [string]$WorkerLabel = ''
    )
    if (-not $Decision) { return $true }
    if ([bool]$Decision.pauseNewClaims) { return $false }
    $slots = 1
    if ($Decision.ContainsKey('recommendedSlots') -and $null -ne $Decision.recommendedSlots) {
        try { $slots = [int]$Decision.recommendedSlots } catch { $slots = 1 }
    }
    if ($slots -le 0) { return $false }
    if (-not $Decision.ContainsKey('claimAllowLabels') -or $null -eq $Decision.claimAllowLabels) {
        return $true
    }
    $allowed = @($Decision.claimAllowLabels | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($allowed.Count -eq 0) { return $true }
    if ([string]::IsNullOrWhiteSpace($WorkerLabel)) { return $true }
    return ($allowed -contains [string]$WorkerLabel)
}

function Get-QCHostThrottleSpawnPlan {
    <#
    .SYNOPSIS
    Decide whether to spawn. Never requests killing existing workers when slots drop.
    #>
    [CmdletBinding()]
    param(
        [int]$CurrentCount,
        [int]$Want
    )
    $cur = $CurrentCount
    if ($cur -lt 0) { $cur = 0 }
    $w = $Want
    if ($w -lt 0) { $w = 0 }
    return @{
        want            = $w
        currentCount    = $cur
        shouldSpawn     = ($cur -lt $w)
        shouldStopExcess = $false
    }
}

function ConvertTo-QCHostThrottleEvaluation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Settings,
        [int]$MaxParallel = 1,
        [AllowNull()][Nullable[double]]$CpuPercent = $null,
        [AllowNull()][Nullable[double]]$MemoryPercent = $null,
        [object[]]$MatchedProcesses = @(),
        [AllowNull()][Nullable[double]]$MatchedProcessCpuPercent = $null,
        [AllowNull()][Nullable[double]]$MatchedProcessMemoryMb = $null,
        [switch]$MatchedProcessUsageReady,
        [switch]$SampleError
    )
    $enabled = $false
    try { $enabled = [bool]$Settings.enabled } catch { $enabled = $false }
    $max = $MaxParallel
    if ($max -lt 1) { $max = 1 }
    $busySlots = 1
    if ($Settings.ContainsKey('busyRecommendedSlots') -and $null -ne $Settings.busyRecommendedSlots) {
        try { $busySlots = [int]$Settings.busyRecommendedSlots } catch { $busySlots = 1 }
    }

    if ($SampleError.IsPresent) {
        return @{
            enabled           = $enabled
            pauseNewClaims    = $false
            recommendedSlots  = $max
            maxParallel       = $max
            reason            = 'sample_error'
            matchedProcesses  = @()
            cpuPercent        = $CpuPercent
            memoryPercent     = $MemoryPercent
            matchedProcessCpuPercent = $MatchedProcessCpuPercent
            matchedProcessMemoryMb   = $MatchedProcessMemoryMb
        }
    }

    if (-not $enabled) {
        return @{
            enabled           = $false
            pauseNewClaims    = $false
            recommendedSlots  = $max
            maxParallel       = $max
            reason            = 'disabled'
            matchedProcesses  = @()
            cpuPercent        = $CpuPercent
            memoryPercent     = $MemoryPercent
            matchedProcessCpuPercent = $MatchedProcessCpuPercent
            matchedProcessMemoryMb   = $MatchedProcessMemoryMb
        }
    }

    $cpuThresh = 0
    $memThresh = 0
    if ($Settings.ContainsKey('cpuPercent') -and $null -ne $Settings.cpuPercent) {
        try { $cpuThresh = [double]$Settings.cpuPercent } catch { $cpuThresh = 0 }
    }
    if ($Settings.ContainsKey('memoryPercent') -and $null -ne $Settings.memoryPercent) {
        try { $memThresh = [double]$Settings.memoryPercent } catch { $memThresh = 0 }
    }

    $matchedNames = @(
        @($MatchedProcesses) | ForEach-Object {
            if ($null -eq $_) { return $null }
            if ($_ -is [string]) { return $_ }
            try { return [string]$_.Name } catch { return $null }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    $procCpuThresh = 0
    $procMemThresh = 0
    if ($Settings.ContainsKey('processCpuPercent') -and $null -ne $Settings.processCpuPercent) {
        try { $procCpuThresh = [double]$Settings.processCpuPercent } catch { $procCpuThresh = 0 }
    }
    if ($Settings.ContainsKey('processMemoryMb') -and $null -ne $Settings.processMemoryMb) {
        try { $procMemThresh = [double]$Settings.processMemoryMb } catch { $procMemThresh = 0 }
    }
    $resourceGated = ($procCpuThresh -gt 0 -or $procMemThresh -gt 0)

    $hasMatch = ($matchedNames.Count -gt 0)
    $procCpuBusy = $false
    $procMemBusy = $false
    if ($resourceGated) {
        if ($hasMatch -and $MatchedProcessUsageReady.IsPresent) {
            $procCpuBusy = ($procCpuThresh -gt 0 -and $null -ne $MatchedProcessCpuPercent -and [double]$MatchedProcessCpuPercent -ge $procCpuThresh)
            $procMemBusy = ($procMemThresh -gt 0 -and $null -ne $MatchedProcessMemoryMb -and [double]$MatchedProcessMemoryMb -ge $procMemThresh)
        }
        $procBusy = ($procCpuBusy -or $procMemBusy)
    } else {
        $procBusy = $hasMatch
    }
    $cpuBusy = ($cpuThresh -gt 0 -and $null -ne $CpuPercent -and [double]$CpuPercent -ge $cpuThresh)
    $memBusy = ($memThresh -gt 0 -and $null -ne $MemoryPercent -and [double]$MemoryPercent -ge $memThresh)
    $signalCount = @($procBusy, $cpuBusy, $memBusy) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count

    $reason = 'normal'
    $pause = $false
    $slots = $max
    if ($signalCount -gt 0) {
        $slots = Get-QCHostThrottleDesiredSlots -MaxParallel $max -RecommendedSlots $busySlots
        $pause = ($slots -le 0)
        if ($signalCount -gt 1) { $reason = 'multiple_signals' }
        elseif ($procBusy -and $resourceGated -and $procCpuBusy -and -not $procMemBusy) { $reason = 'process_cpu_threshold' }
        elseif ($procBusy -and $resourceGated -and $procMemBusy -and -not $procCpuBusy) { $reason = 'process_memory_threshold' }
        elseif ($procBusy -and $resourceGated) { $reason = 'process_cpu_threshold' }
        elseif ($procBusy) { $reason = 'matched_process' }
        elseif ($cpuBusy) { $reason = 'cpu_threshold' }
        else { $reason = 'memory_threshold' }
    }

    return @{
        enabled           = $true
        pauseNewClaims    = $pause
        recommendedSlots  = $slots
        maxParallel       = $max
        reason            = $reason
        matchedProcesses  = @($matchedNames)
        cpuPercent        = $CpuPercent
        memoryPercent     = $MemoryPercent
        matchedProcessCpuPercent = $MatchedProcessCpuPercent
        matchedProcessMemoryMb   = $MatchedProcessMemoryMb
    }
}

function Get-QCHostThrottleHealth {
    <#
    .SYNOPSIS
    Classify a throttle heartbeat for ops display: healthy, throttled, stale, disabled.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Status,
        [int]$SampleSeconds = 10,
        [datetime]$NowUtc = [datetime]::MinValue
    )
    if (-not $Status) { return 'stale' }
    $reason = ''
    try { $reason = [string]$Status.reason } catch { $reason = '' }
    $enabled = $true
    if ($Status.ContainsKey('enabled') -and $null -ne $Status.enabled) {
        try { $enabled = [bool]$Status.enabled } catch { $enabled = $true }
    }
    $fresh = Test-QCHostThrottleStatusFresh -Status $Status -SampleSeconds $SampleSeconds -NowUtc $NowUtc
    if (-not $fresh) { return 'stale' }
    if ((-not $enabled) -or ($reason -eq 'disabled')) { return 'disabled' }
    $pause = $false
    if ($Status.ContainsKey('pauseNewClaims') -and $null -ne $Status.pauseNewClaims) {
        try { $pause = [bool]$Status.pauseNewClaims } catch { $pause = $false }
    }
    $slots = $null
    if ($Status.ContainsKey('recommendedSlots') -and $null -ne $Status.recommendedSlots) {
        try { $slots = [int]$Status.recommendedSlots } catch { $slots = $null }
    }
    $max = $null
    if ($Status.ContainsKey('maxParallel') -and $null -ne $Status.maxParallel) {
        try { $max = [int]$Status.maxParallel } catch { $max = $null }
    }
    if ($reason -eq 'sample_error') { return 'throttled' }
    if ($pause) { return 'throttled' }
    if ($null -ne $slots -and $slots -le 0) { return 'throttled' }
    if ($null -ne $slots -and $null -ne $max -and $slots -lt $max) { return 'throttled' }
    return 'healthy'
}

function Get-QCHostThrottleQueueRoot {
    [CmdletBinding()]
    param([hashtable]$Config)
    $root = $null
    try {
        if ($Config -and $Config.ContainsKey('queue') -and $Config.queue) {
            if ($Config.queue.rootDir) { $root = [string]$Config.queue.rootDir }
            elseif ($Config.queue.root) { $root = [string]$Config.queue.root }
        }
    } catch { $root = $null }
    return $root
}

function Get-QCHostThrottleStatusPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [string]$HostName = ''
    )
    $hostVal = $HostName
    if ([string]::IsNullOrWhiteSpace($hostVal)) { $hostVal = [string]$env:COMPUTERNAME }
    $safe = ($hostVal -replace '[^A-Za-z0-9._-]', '_')
    return (Join-Path $QueueRoot ('_remote_worker.{0}.throttle.json' -f $safe))
}

function ConvertTo-QCHostThrottleStatusHashtable {
    param([object]$Obj)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [hashtable]) { return $Obj }
    $h = @{}
    foreach ($p in @($Obj.PSObject.Properties)) {
        $h[$p.Name] = $p.Value
    }
    return $h
}

function Write-QCHostThrottleStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Evaluation,
        [string]$HostName = '',
        [datetime]$SampledAtUtc = [datetime]::MinValue
    )
    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $when = $SampledAtUtc
    if ($when -eq [datetime]::MinValue) { $when = [datetime]::UtcNow }
    if ($when.Kind -eq [DateTimeKind]::Local) { $when = $when.ToUniversalTime() }
    $hostVal = $HostName
    if ([string]::IsNullOrWhiteSpace($hostVal)) { $hostVal = [string]$env:COMPUTERNAME }

    $cpu = $null
    if ($Evaluation.ContainsKey('cpuPercent') -and $null -ne $Evaluation.cpuPercent) {
        try { $cpu = [math]::Round([double]$Evaluation.cpuPercent, 1) } catch { $cpu = $null }
    }
    $mem = $null
    if ($Evaluation.ContainsKey('memoryPercent') -and $null -ne $Evaluation.memoryPercent) {
        try { $mem = [math]::Round([double]$Evaluation.memoryPercent, 1) } catch { $mem = $null }
    }
    $matchedProcCpu = $null
    if ($Evaluation.ContainsKey('matchedProcessCpuPercent') -and $null -ne $Evaluation.matchedProcessCpuPercent) {
        try { $matchedProcCpu = [math]::Round([double]$Evaluation.matchedProcessCpuPercent, 1) } catch { $matchedProcCpu = $null }
    }
    $matchedProcMb = $null
    if ($Evaluation.ContainsKey('matchedProcessMemoryMb') -and $null -ne $Evaluation.matchedProcessMemoryMb) {
        try { $matchedProcMb = [math]::Round([double]$Evaluation.matchedProcessMemoryMb, 1) } catch { $matchedProcMb = $null }
    }

    $allow = @()
    if ($Evaluation.ContainsKey('claimAllowLabels') -and $null -ne $Evaluation.claimAllowLabels) {
        $allow = @($Evaluation.claimAllowLabels | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    $payload = [ordered]@{
        enabled           = [bool]$Evaluation.enabled
        pauseNewClaims    = [bool]$Evaluation.pauseNewClaims
        recommendedSlots  = [int]$Evaluation.recommendedSlots
        reason            = [string]$Evaluation.reason
        matchedProcesses  = @($Evaluation.matchedProcesses)
        matchedProcessCpuPercent = $matchedProcCpu
        matchedProcessMemoryMb   = $matchedProcMb
        claimAllowLabels  = @($allow)
        cpuPercent        = $cpu
        memoryPercent     = $mem
        sampledAtUtc      = $when.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        host              = $hostVal
    }
    if ($Evaluation.ContainsKey('maxParallel') -and $null -ne $Evaluation.maxParallel) {
        try { $payload['maxParallel'] = [int]$Evaluation.maxParallel } catch { }
    }
    $tmp = $Path + '.tmp.' + ([guid]::NewGuid().ToString('N'))
    $enc = New-Object System.Text.UTF8Encoding $false
    try {
        $json = ($payload | ConvertTo-Json -Compress -Depth 6)
        [System.IO.File]::WriteAllText($tmp, $json, $enc)
        Move-Item -LiteralPath $tmp -Destination $Path -Force -ErrorAction Stop
    } catch {
        try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } } catch { }
        throw
    }
    return $payload
}

function Read-QCHostThrottleStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    try {
        $raw = [System.IO.File]::ReadAllText($Path)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        return (ConvertTo-QCHostThrottleStatusHashtable -Obj $obj)
    } catch {
        return $null
    }
}

function Test-QCHostThrottleStatusFresh {
    [CmdletBinding()]
    param(
        [hashtable]$Status,
        [int]$SampleSeconds = 10,
        [datetime]$NowUtc = [datetime]::MinValue
    )
    if (-not $Status -or -not $Status.ContainsKey('sampledAtUtc') -or [string]::IsNullOrWhiteSpace([string]$Status.sampledAtUtc)) {
        return $false
    }
    $now = $NowUtc
    if ($now -eq [datetime]::MinValue) { $now = [datetime]::UtcNow }
    $now = $now.ToUniversalTime()
    $sampled = $null
    try {
        $sampled = [datetimeoffset]::Parse([string]$Status.sampledAtUtc, [System.Globalization.CultureInfo]::InvariantCulture).UtcDateTime
    } catch {
        try { $sampled = [datetime]::Parse([string]$Status.sampledAtUtc, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime() } catch { return $false }
    }
    $freshFor = Get-QCHostThrottleFreshnessSeconds -SampleSeconds $SampleSeconds
    return (($now - $sampled).TotalSeconds -le $freshFor)
}

function Get-QCHostThrottleClaimDecision {
    <#
    .SYNOPSIS
    Fail-open claim gate. Pause only when a fresh valid status explicitly says to.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Status,
        [hashtable]$Settings,
        [int]$MaxParallel = 1,
        [datetime]$NowUtc = [datetime]::MinValue
    )
    $max = $MaxParallel
    if ($max -lt 1) { $max = 1 }
    $cfgEnabled = $false
    $sampleSeconds = 10
    if ($Settings) {
        try { $cfgEnabled = [bool]$Settings.enabled } catch { $cfgEnabled = $false }
        if ($Settings.ContainsKey('sampleSeconds')) {
            try { $sampleSeconds = [int]$Settings.sampleSeconds } catch { $sampleSeconds = 10 }
        }
    }
    $open = @{
        pauseNewClaims   = $false
        recommendedSlots = $max
        reason           = 'disabled'
        honored          = $false
        fresh            = $false
        claimAllowLabels = @()
    }
    if (-not $cfgEnabled) { return $open }

    if (-not $Status) {
        $open.reason = 'normal'
        return $open
    }

    $reason = ''
    try { $reason = [string]$Status.reason } catch { $reason = '' }
    if ($reason -eq 'sample_error' -or $reason -eq 'disabled') {
        $open.reason = $(if ($reason) { $reason } else { 'normal' })
        return $open
    }

    $fresh = Test-QCHostThrottleStatusFresh -Status $Status -SampleSeconds $sampleSeconds -NowUtc $NowUtc
    $open.fresh = $fresh
    if (-not $fresh) {
        $open.reason = 'normal'
        return $open
    }

    $statusEnabled = $true
    if ($Status.ContainsKey('enabled')) {
        try { $statusEnabled = [bool]$Status.enabled } catch { $statusEnabled = $true }
    }
    if (-not $statusEnabled) {
        $open.reason = 'disabled'
        $open.honored = $true
        $open.fresh = $true
        return $open
    }

    $pause = $false
    if ($Status.ContainsKey('pauseNewClaims')) {
        try { $pause = [bool]$Status.pauseNewClaims } catch { $pause = $false }
    }
    $slots = $max
    if ($Status.ContainsKey('recommendedSlots') -and $null -ne $Status.recommendedSlots) {
        try { $slots = [int]$Status.recommendedSlots } catch { $slots = $max }
    }
    $slots = Get-QCHostThrottleDesiredSlots -MaxParallel $max -RecommendedSlots $slots
    $allow = @()
    if ($Status.ContainsKey('claimAllowLabels') -and $null -ne $Status.claimAllowLabels) {
        $allow = @($Status.claimAllowLabels | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if (-not $reason) { $reason = $(if ($pause) { 'matched_process' } else { 'normal' }) }
    if ($pause -or $slots -le 0) { $pause = $true; $slots = 0 }
    return @{
        pauseNewClaims   = [bool]$pause
        recommendedSlots = $slots
        reason           = $reason
        honored          = $true
        fresh            = $true
        claimAllowLabels = @($allow)
    }
}

function Get-QCHostCpuPercentLive {
    $rows = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop)
    if ($rows.Count -eq 0) { return $null }
    return [double](($rows | Measure-Object -Property LoadPercentage -Average).Average)
}

function Get-QCHostMemoryPercentLive {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $total = [double]$os.TotalVisibleMemorySize
    $free = [double]$os.FreePhysicalMemory
    if ($total -le 0) { return $null }
    return [math]::Round((($total - $free) / $total) * 100.0, 1)
}

function Get-QCHostProcessSnapshotLive {
    @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | ForEach-Object {
        [pscustomobject]@{
            ProcessId       = [int]$_.ProcessId
            Name            = [string]$_.Name
            ParentProcessId = [int]$_.ParentProcessId
            KernelModeTime  = $_.KernelModeTime
            UserModeTime    = $_.UserModeTime
            WorkingSetSize  = $_.WorkingSetSize
        }
    })
}

function Get-QCHostResourceSample {
    <#
    .SYNOPSIS
    Compose a resource sample. Pass CpuPercent/MemoryPercent/Processes to skip live CIM.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][Nullable[double]]$CpuPercent,
        [AllowNull()][Nullable[double]]$MemoryPercent,
        [object[]]$Processes,
        [switch]$Live
    )
    $cpu = $CpuPercent
    $mem = $MemoryPercent
    $procs = $Processes
    if ($Live.IsPresent) {
        if ($null -eq $cpu) { $cpu = Get-QCHostCpuPercentLive }
        if ($null -eq $mem) { $mem = Get-QCHostMemoryPercentLive }
        if ($null -eq $procs) { $procs = Get-QCHostProcessSnapshotLive }
    }
    return @{
        cpuPercent    = $cpu
        memoryPercent = $mem
        processes     = @($procs)
        sampledAtUtc  = [datetime]::UtcNow
    }
}

function Update-QCHostThrottleStatus {
    <#
    .SYNOPSIS
    Sample (or use injected sample), evaluate, and write the host throttle status file.
    Sampling exceptions write reason=sample_error and do not pause claims.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Settings,
        [Parameter(Mandatory)][string]$StatusPath,
        [int]$MaxParallel = 1,
        [int[]]$ExcludePids = @(),
        [hashtable]$Sample = $null,
        [string]$HostName = '',
        [string[]]$CurrentLabels = @(),
        [object[]]$PreviousProcesses,
        [datetime]$PreviousSampledAt = [datetime]::MinValue,
        [switch]$Live
    )
    $cpu = $null
    $mem = $null
    $matched = @()
    try {
        $snap = $Sample
        if (-not $snap) {
            $snap = Get-QCHostResourceSample -Live:$Live.IsPresent
        }
        if ($snap.ContainsKey('cpuPercent')) { $cpu = $snap.cpuPercent }
        if ($snap.ContainsKey('memoryPercent')) { $mem = $snap.memoryPercent }
        $procs = @()
        if ($snap.ContainsKey('processes')) { $procs = @($snap.processes) }
        $pats = @()
        if ($Settings.ContainsKey('processNamePatterns')) { $pats = @($Settings.processNamePatterns) }
        $tree = Get-QCHostMatchedProcessTreePids -Processes $procs -Patterns $pats -ExcludePids $ExcludePids
        $matched = @($tree.matched)
        $prev = $script:QCHostThrottlePrevProcesses
        $prevAt = $script:QCHostThrottlePrevAt
        if ($PSBoundParameters.ContainsKey('PreviousProcesses')) { $prev = $PreviousProcesses }
        if ($PreviousSampledAt -ne [datetime]::MinValue) { $prevAt = $PreviousSampledAt }
        $whenSample = [datetime]::UtcNow
        if ($snap.ContainsKey('sampledAtUtc') -and $snap.sampledAtUtc) {
            try { $whenSample = [datetime]$snap.sampledAtUtc } catch { $whenSample = [datetime]::UtcNow }
        }
        $elapsed = 0.0
        if ($null -ne $prevAt) {
            try { $elapsed = ($whenSample.ToUniversalTime() - ([datetime]$prevAt).ToUniversalTime()).TotalSeconds } catch { $elapsed = 0.0 }
        }
        $usage = Measure-QCHostMatchedProcessUsage -PreviousProcesses $prev -CurrentProcesses $procs `
            -CandidatePids $tree.treePids -ElapsedSeconds $elapsed
        $evalParams = @{
            Settings = $Settings
            MaxParallel = $MaxParallel
            CpuPercent = $cpu
            MemoryPercent = $mem
            MatchedProcesses = $matched
            MatchedProcessCpuPercent = $usage.sumCpuPercent
            MatchedProcessMemoryMb = $usage.sumWorkingSetMb
        }
        if ($usage.comparable) { $evalParams['MatchedProcessUsageReady'] = $true }
        $eval = ConvertTo-QCHostThrottleEvaluation @evalParams
        $script:QCHostThrottlePrevProcesses = $procs
        $script:QCHostThrottlePrevAt = $whenSample
    } catch {
        $eval = ConvertTo-QCHostThrottleEvaluation -Settings $Settings -MaxParallel $MaxParallel `
            -CpuPercent $cpu -MemoryPercent $mem -MatchedProcesses @() -SampleError
        $eval['sampleError'] = [string]$_.Exception.Message
    }
    $want = Get-QCHostThrottleDesiredSlots -MaxParallel $MaxParallel -RecommendedSlots ([int]$eval.recommendedSlots)
    $eval['claimAllowLabels'] = @(Get-QCHostThrottleClaimAllowLabels -CurrentLabels $CurrentLabels -RecommendedSlots $want)
    $when = [datetime]::UtcNow
    if ($Sample -and $Sample.ContainsKey('sampledAtUtc') -and $Sample.sampledAtUtc) {
        try { $when = [datetime]$Sample.sampledAtUtc } catch { $when = [datetime]::UtcNow }
    }
    $written = Write-QCHostThrottleStatus -Path $StatusPath -Evaluation $eval -HostName $HostName -SampledAtUtc $when
    return @{
        evaluation = $eval
        status     = $written
        path       = $StatusPath
    }
}

Export-ModuleMember -Function *
