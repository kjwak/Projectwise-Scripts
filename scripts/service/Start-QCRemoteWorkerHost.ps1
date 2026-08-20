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
                } catch {
                    Write-QCJsonLog -Level 'Error' -Code 'REMOTE_HOST_SPAWN_FAILED' -Message $_.Exception.Message -Data @{ label = $label }
                }
            }
        }

        if (([DateTime]::UtcNow - $lastStatusAt).TotalSeconds -ge 10) {
            $lastStatusAt = [DateTime]::UtcNow
            $pids = @()
            foreach ($lbl in @($script:WorkerSlots.Keys)) {
                try { $pids += [int]$script:WorkerSlots[$lbl].process.Id } catch { }
            }
            $types = if (@($rh.enabledJobTypes).Count -gt 0) { ($rh.enabledJobTypes -join ',') } else { '(all)' }
            Write-Host ("[{0}] remote-host {1} workers={2}/{3} types={4} pids={5}" -f (Get-Date -Format 'HH:mm:ss'), $hostName, $script:WorkerSlots.Count, $want, $types, ($pids -join ','))
        }

        Start-Sleep -Milliseconds 400
    }
} finally {
    & $exitHandler
}
