<#
.SYNOPSIS
Processor-only supervisor for a remote QC worker host. Spawns Run-QCProcessor.ps1
children. Does not start the watcher or dashboard.

.DESCRIPTION
Use this on the modelling PC. The QC server remains the coordinator (watcher +
JSON queue). This host only claims jobs.

UNC queue roots are refused unless you pass -AllowUncQueue or set
workers.remoteHost.allowUncQueue. Restrict types with workers.enabledJobTypes
(empty/missing = all types; this PC typically uses ["QC_PREPEND"]).

.EXAMPLE
.\scripts\service\Start-QCRemoteWorkerHost.ps1 -AllowUncQueue
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$AllowUncQueue
)

$ErrorActionPreference = 'Stop'

$scriptsRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $scriptsRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -RepoRoot $repoRoot -FeatureModules @(
    'Core\Core.Results.psm1'
    'Queue\QC.Queue.Json.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Get-QCRemoteWorkerHostSettings'
    'Test-QCUncQueueClaimAllowed'
    'Recover-QCStaleJobs'
    'Write-QCJsonLog'
    'Get-QCTimestamp'
    'Get-QCLogHourStamp'
) -Context 'remote worker host bootstrap'

function _CmdLineEscapeDoubleQuotes([string]$Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace('"', '""')
}

function _Append-CmdLineArg([System.Text.StringBuilder]$Sb, [string]$Value) {
    if ($null -eq $Value) { return }
    $s = [string]$Value
    if ($s -match '[\s"]') {
        [void]$Sb.Append('"')
        [void]$Sb.Append((_CmdLineEscapeDoubleQuotes $s))
        [void]$Sb.Append('"')
    } else {
        [void]$Sb.Append($s)
    }
}

function _Get-QueueRoot([hashtable]$Cfg) {
    $root = $null
    try {
        if ($Cfg.queue -and $Cfg.queue.rootDir) { $root = [string]$Cfg.queue.rootDir }
        elseif ($Cfg.queue -and $Cfg.queue.root) { $root = [string]$Cfg.queue.root }
    } catch { }
    if (-not $root) { $root = Join-Path $repoRoot 'queue' }
    return $root
}

function _Get-ChildLogDir([string]$QueueRoot) {
    $logDir = Join-Path $QueueRoot '_logs'
    if (-not (Test-Path -LiteralPath $logDir)) {
        try { New-Item -ItemType Directory -Path $logDir -Force | Out-Null } catch { }
    }
    return $logDir
}

function _Start-WorkerProcess {
    param(
        [string]$ScriptPath,
        [string[]]$ScriptArgs,
        [string]$LogDir
    )
    $tag = 'Run-QCProcessor'
    $stamp = Get-QCTimestamp
    $stampShort = $stamp -replace '[:.]', ''
    $stdoutPath = Join-Path $LogDir ("${stampShort}_${tag}.discard.out")
    $stderrPath = Join-Path $LogDir ("${stampShort}_${tag}.err.log")
    $savedLogDir = $env:QC_JSON_LOG_DIR
    $savedLogTag = $env:QC_JSON_LOG_TAG
    $env:QC_JSON_LOG_DIR = $LogDir
    $env:QC_JSON_LOG_TAG = $tag
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('-NoProfile -ExecutionPolicy Bypass -MTA -File ')
    _Append-CmdLineArg -Sb $sb -Value $ScriptPath
    foreach ($a in @($ScriptArgs)) {
        [void]$sb.Append(' ')
        _Append-CmdLineArg -Sb $sb -Value ([string]$a)
    }
    try {
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $sb.ToString() -NoNewWindow -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    } finally {
        if ($null -ne $savedLogDir) { $env:QC_JSON_LOG_DIR = $savedLogDir } else { Remove-Item -Path 'Env:QC_JSON_LOG_DIR' -ErrorAction SilentlyContinue }
        if ($null -ne $savedLogTag) { $env:QC_JSON_LOG_TAG = $savedLogTag } else { Remove-Item -Path 'Env:QC_JSON_LOG_TAG' -ErrorAction SilentlyContinue }
    }
    return @{
        process = $p
        stdoutPath = $stdoutPath
        stderrPath = $stderrPath
    }
}

function _Get-ProcessorJsonlPath([string]$LogDir, [string]$HourStamp) {
    return (Join-Path $LogDir ("Run-QCProcessor_${HourStamp}.jsonl"))
}

function _Read-LogChunkFromOffset([string]$Path, [int64]$StartPos) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        return @{ Text = ''; NewPos = [int64]$StartPos }
    }
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            if ($StartPos -gt $fs.Length) { $StartPos = [int64]0 }
            $fs.Position = $StartPos
            $sr = New-Object System.IO.StreamReader($fs, [System.Text.UTF8Encoding]::new($false), $false, 4096, $true)
            $text = $sr.ReadToEnd()
            $newPos = $fs.Position
            $sr.Dispose()
            return @{ Text = [string]$text; NewPos = [int64]$newPos }
        } finally {
            $fs.Dispose()
        }
    } catch {
        return @{ Text = ''; NewPos = [int64]$StartPos }
    }
}

function _Test-RemoteHostJobEvent([string]$Code) {
    if ([string]::IsNullOrWhiteSpace($Code)) { return $false }
    if ($Code -in @(
            'WORKER_NO_JOB'
            'WORKER_START'
            'WORKER_BUDGET'
            'WORKER_LEASE'
            'WORKER_LOCK_RACE'
            'WORKER_DB_SCHEMA_INIT_FAILED'
            'WORKER_QUEUE_DUPLICATE_CLEANUP'
            'JOB_TELEMETRY_WRITTEN'
            'JOB_TELEMETRY_SKIPPED'
        )) { return $false }
    if ($Code -match '^(WORKER_|QC_PREPEND|STATUS_SET)') { return $true }
    return $false
}

function _Write-RemoteHostJobLine {
    param([object]$Obj)
    if (-not $Obj) { return }
    $code = [string]$Obj.code
    if (-not (_Test-RemoteHostJobEvent $code)) { return }

    $d = $null
    try { $d = $Obj.data } catch { }
    $label = ''; $jobId = ''; $jobType = ''; $src = ''; $stage = ''
    if ($d) {
        try { if ($d.workerLabel) { $label = [string]$d.workerLabel } } catch { }
        try { if ($d.jobId) { $jobId = [string]$d.jobId } } catch { }
        try { if ($d.jobType) { $jobType = [string]$d.jobType } } catch { }
        try {
            if ($d.sourceName) { $src = [string]$d.sourceName }
            elseif ($d.sourcePath) { $src = [System.IO.Path]::GetFileName([string]$d.sourcePath) }
        } catch { }
        try { if ($d.stage) { $stage = [string]$d.stage } } catch { }
    }
    if (-not $src -and $jobId -and $script:ActiveJobs -and $script:ActiveJobs.ContainsKey($jobId)) {
        try { $src = [string]$script:ActiveJobs[$jobId] } catch { }
    }
    if ($code -eq 'WORKER_STAGE' -and -not $jobId) { return }

    $msg = [string]$Obj.message
    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add(('[{0}]' -f (Get-Date -Format 'HH:mm:ss')))
    if ($label) { [void]$parts.Add($label) }
    [void]$parts.Add($code)
    if ($jobType) { [void]$parts.Add($jobType) }
    if ($src) { [void]$parts.Add($src) }
    if ($jobId) { [void]$parts.Add($jobId) }
    if ($stage) { [void]$parts.Add(("stage={0}" -f $stage)) }
    if ($msg) { [void]$parts.Add($msg) }
    $line = ($parts -join ' ')

    $color = 'White'
    switch -Regex ($code) {
        '^(WORKER_SELECTED|WORKER_CLAIMING)$' { $color = 'Cyan' }
        '^WORKER_STAGE$' { $color = 'Yellow' }
        '^(WORKER_SUCCEEDED|QC_PREPEND_OK)$' { $color = 'Green' }
        '(FAILED|UNHANDLED|ERROR)' { $color = 'Red' }
        default { $color = 'Gray' }
    }
    Write-Host $line -ForegroundColor $color

    if ($jobId) {
        if ($code -match '^(WORKER_SELECTED|WORKER_CLAIMING)$') {
            $script:ActiveJobs[$jobId] = $(if ($src) { $src } else { $jobType })
        } elseif ($code -match '^(WORKER_SUCCEEDED|WORKER_FAILED|WORKER_JOB_UNHANDLED)$') {
            try { $script:ActiveJobs.Remove($jobId) | Out-Null } catch { }
        }
    }
}

function _Drain-ProcessorJsonLogs([string]$LogDir) {
    $hour = Get-QCLogHourStamp
    if ($script:ProcJsonLog.hour -and [string]$script:ProcJsonLog.hour -ne $hour) {
        $prev = _Get-ProcessorJsonlPath -LogDir $LogDir -HourStamp ([string]$script:ProcJsonLog.hour)
        _Drain-ProcessorJsonlFile -Path $prev
        $script:ProcJsonLog.offset = [int64]0
        $script:ProcJsonLog.tail = ''
        $script:ProcJsonLog.skipExisting = $false
    }
    $script:ProcJsonLog.hour = $hour
    $path = _Get-ProcessorJsonlPath -LogDir $LogDir -HourStamp $hour
    if ($script:ProcJsonLog.skipExisting) {
        if (Test-Path -LiteralPath $path) {
            try { $script:ProcJsonLog.offset = [int64](Get-Item -LiteralPath $path).Length } catch { $script:ProcJsonLog.offset = [int64]0 }
        }
        $script:ProcJsonLog.skipExisting = $false
        return
    }
    _Drain-ProcessorJsonlFile -Path $path
}

function _Drain-ProcessorJsonlFile([string]$Path) {
    $chunk = _Read-LogChunkFromOffset -Path $Path -StartPos ([int64]$script:ProcJsonLog.offset)
    if ($chunk.NewPos -lt $script:ProcJsonLog.offset) {
        $script:ProcJsonLog.offset = [int64]0
        $script:ProcJsonLog.tail = ''
        $chunk = _Read-LogChunkFromOffset -Path $Path -StartPos 0
    }
    $script:ProcJsonLog.offset = [int64]$chunk.NewPos
    $buffer = ([string]$script:ProcJsonLog.tail) + ([string]$chunk.Text)
    $nl = $buffer.LastIndexOfAny(@([char]10, [char]13))
    if ($nl -lt 0) {
        $script:ProcJsonLog.tail = $buffer
        return
    }
    $complete = $buffer.Substring(0, $nl + 1)
    $script:ProcJsonLog.tail = $buffer.Substring($nl + 1)
    foreach ($line in ($complete -split '\r?\n')) {
        $t = ([string]$line).Trim()
        if (-not $t -or -not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            _Write-RemoteHostJobLine -Obj $o
        } catch { }
    }
}

$cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath -DryRun:$DryRun.IsPresent
if (-not (Test-QCUncQueueClaimAllowed -Config $cfg -AllowUncQueue:$AllowUncQueue.IsPresent)) {
    throw "Queue root is a UNC path. Pass -AllowUncQueue or set workers.remoteHost.allowUncQueue before claiming jobs from this host."
}

$rh = Get-QCRemoteWorkerHostSettings -Config $cfg
$queueRoot = _Get-QueueRoot -Cfg $cfg
$logDir = _Get-ChildLogDir -QueueRoot $queueRoot
$workerScript = Join-Path $PSScriptRoot 'Run-QCProcessor.ps1'
$hostName = [string]$env:COMPUTERNAME
$safeHost = ($hostName -replace '[^A-Za-z0-9._-]', '_')
$lockPath = Join-Path $queueRoot ("_remote_worker.$safeHost.lock")

if (Test-Path -LiteralPath $lockPath) {
    $existing = $null
    try { $existing = Get-Content -LiteralPath $lockPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch { }
    $ownerPid = 0
    try { if ($existing -and $existing.pid) { $ownerPid = [int]$existing.pid } } catch { $ownerPid = 0 }
    if ($ownerPid -gt 0) {
        $proc = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
        $isPw = $false
        try { if ($proc) { $isPw = ($proc.ProcessName -in @('powershell', 'pwsh')) } } catch { $isPw = $false }
        if ($proc -and $isPw) {
            throw "Remote worker host already running on $hostName (pid=$ownerPid). Stop it first with .\scripts\service\Stop-QCPipeline.ps1"
        }
    }
    try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue } catch { }
}

@{
    pid          = $PID
    startedAtUtc = Get-QCTimestamp
    host         = $hostName
    scriptPath   = $MyInvocation.MyCommand.Path
    queueRoot    = $queueRoot
    enabledJobTypes = @($rh.enabledJobTypes)
} | ConvertTo-Json -Compress | Set-Content -LiteralPath $lockPath -Encoding utf8 -Force

$script:HostLockPath = $lockPath
$script:WorkerSlots = @{}
$script:ActiveJobs = @{}
$script:ProcJsonLog = @{
    hour = ''
    offset = [int64]0
    tail = ''
    skipExisting = $true
}

function _Stop-TrackedWorkers {
    foreach ($lbl in @($script:WorkerSlots.Keys)) {
        $child = $script:WorkerSlots[$lbl]
        try {
            if ($child.process -and -not $child.process.HasExited) {
                Stop-Process -Id $child.process.Id -Force -ErrorAction Stop
            }
        } catch { }
    }
    $script:WorkerSlots = @{}
}

$exitHandler = {
    try { _Stop-TrackedWorkers } catch { }
    try {
        if ($script:HostLockPath -and (Test-Path -LiteralPath $script:HostLockPath)) {
            Remove-Item -LiteralPath $script:HostLockPath -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}
[System.AppDomain]::CurrentDomain.add_ProcessExit($exitHandler)

Write-QCJsonLog -Level 'Information' -Code 'REMOTE_HOST_START' -Message 'Remote worker host started (processor-only; no watcher).' -Data @{
    host = $hostName
    queueRoot = $queueRoot
    maxParallel = [int]$rh.maxParallel
    enabledJobTypes = @($rh.enabledJobTypes)
    allowUncQueue = [bool]($AllowUncQueue.IsPresent -or $rh.allowUncQueue)
    dryRun = [bool]$cfg.dryRun
}

Write-Host ("[{0}] remote-host started host={1} queue={2} types={3}" -f (Get-Date -Format 'HH:mm:ss'), $hostName, $queueRoot, $(if (@($rh.enabledJobTypes).Count -gt 0) { $rh.enabledJobTypes -join ',' } else { '(all)' })) -ForegroundColor Gray
Write-Host 'Tailing processor JSON logs: claim, stage, success, and failure print here. Idle heartbeat every 10s.' -ForegroundColor Gray

try {
    $rec = Recover-QCStaleJobs -Config $cfg
    if ($rec -and $rec.IsSuccess -and $rec.Data) {
        Write-QCJsonLog -Level 'Information' -Code 'REMOTE_HOST_RECOVERY' -Message 'Startup stale-job recovery.' -Data $rec.Data
    }
} catch { }

$nextWorkerIndex = 1
$lastSpawnAt = [DateTime]::MinValue
$lastRecoveryAt = [DateTime]::UtcNow
$lastStatusAt = [DateTime]::MinValue

try {
    while ($true) {
        $cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath -DryRun:$DryRun.IsPresent
        $rh = Get-QCRemoteWorkerHostSettings -Config $cfg

        $dead = @()
        foreach ($lbl in @($script:WorkerSlots.Keys)) {
            $child = $script:WorkerSlots[$lbl]
            $alive = $false
            try { $alive = -not $child.process.HasExited } catch { $alive = $false }
            if (-not $alive) { $dead += $lbl }
        }
        foreach ($lbl in $dead) { $script:WorkerSlots.Remove($lbl) | Out-Null }

        if (([DateTime]::UtcNow - $lastRecoveryAt).TotalSeconds -ge 30) {
            $lastRecoveryAt = [DateTime]::UtcNow
            try {
                $rec = Recover-QCStaleJobs -Config $cfg
                $req = 0; $fai = 0; $orph = 0
                try { $req = [int]$rec.Data.recoveredToPending } catch { }
                try { $fai = [int]$rec.Data.recoveredToFailed } catch { }
                try { $orph = [int]$rec.Data.recoveredOrphan } catch { }
                if (($req + $fai + $orph) -gt 0) {
                    Write-QCJsonLog -Level 'Information' -Code 'REMOTE_HOST_RECOVERY' -Message 'Periodic stale-job recovery.' -Data $rec.Data
                }
            } catch { }
        }

        $want = [int]$rh.maxParallel
        if ($want -lt 1) { $want = 1 }
        if ($script:WorkerSlots.Count -lt $want) {
            $now = Get-Date
            if (($now - $lastSpawnAt).TotalMilliseconds -ge [int]$rh.spawnStaggerMs) {
                $label = "RW$nextWorkerIndex"
                $nextWorkerIndex++
                $xArgs = @(
                    '-AppSettingsPath', $AppSettingsPath,
                    '-MaxJobs', [string]([int]$rh.maxJobsPerWorker),
                    '-LeaseSeconds', [string]([int]$rh.leaseSeconds),
                    '-IdleSleepMs', [string]([int]$rh.idleSleepMs),
                    '-WorkerLabel', $label
                )
                if ([bool]$cfg.dryRun) { $xArgs += '-DryRun' }
                if ($AllowUncQueue.IsPresent -or [bool]$rh.allowUncQueue) { $xArgs += '-AllowUncQueue' }
                try {
                    $child = _Start-WorkerProcess -ScriptPath $workerScript -ScriptArgs $xArgs -LogDir $logDir
                    $script:WorkerSlots[$label] = $child
                    $lastSpawnAt = $now
                    Write-QCJsonLog -Level 'Information' -Code 'REMOTE_HOST_SPAWN' -Message 'Spawned processor.' -Data @{
                        label = $label
                        pid = [int]$child.process.Id
                    }
                    Write-Host ("[{0}] spawned {1} pid={2}" -f (Get-Date -Format 'HH:mm:ss'), $label, $child.process.Id) -ForegroundColor Gray
                } catch {
                    Write-QCJsonLog -Level 'Error' -Code 'REMOTE_HOST_SPAWN_FAILED' -Message $_.Exception.Message -Data @{ label = $label }
                }
            }
        }

        try { _Drain-ProcessorJsonLogs -LogDir $logDir } catch { }

        if (([DateTime]::UtcNow - $lastStatusAt).TotalSeconds -ge 10) {
            $lastStatusAt = [DateTime]::UtcNow
            $pids = @()
            foreach ($lbl in @($script:WorkerSlots.Keys)) {
                try { $pids += [int]$script:WorkerSlots[$lbl].process.Id } catch { }
            }
            $types = if (@($rh.enabledJobTypes).Count -gt 0) { ($rh.enabledJobTypes -join ',') } else { '(all)' }
            $busyN = 0
            $busyTxt = 'idle'
            try {
                $busyN = @($script:ActiveJobs.Keys).Count
                if ($busyN -gt 0) {
                    $busyTxt = 'busy=' + $busyN + ' ' + ((@($script:ActiveJobs.Values) | Select-Object -First 3) -join ',')
                }
            } catch { }
            Write-Host ("[{0}] remote-host {1} workers={2}/{3} types={4} pids={5} {6}" -f (Get-Date -Format 'HH:mm:ss'), $hostName, $script:WorkerSlots.Count, $want, $types, ($pids -join ','), $busyTxt)
        }

        Start-Sleep -Milliseconds 400
    }
} finally {
    & $exitHandler
}
