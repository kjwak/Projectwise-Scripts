<#
.SYNOPSIS
Live console for the QC pipeline dashboard scheduled task (read-only tail).

.DESCRIPTION
Does not start the dashboard. Shows scheduled-task state, dashboard lock liveness,
queue folder counts, running jobs, watcher JSONL (enqueue/errors), and worker JSONL.
Close the window anytime; Session 0 keeps running.

.EXAMPLE
.\scripts\service\Watch-QCPipelineDashboardConsole.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [string]$ScheduledTaskName = 'QC-PipelineDashboard',

    [Parameter(Mandatory = $false)]
    [int]$RecentLines = 40
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
    'Get-QCLogHourStamp'
) -Context 'pipeline dashboard log console'

. (Join-Path $PSScriptRoot 'QC.RemoteWorkerHostLogView.ps1')

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
    return (Join-Path $QueueRoot '_logs')
}

function _Get-ScheduledTaskStateText {
    param([string]$Name)
    try {
        $t = Get-ScheduledTask -TaskName $Name -ErrorAction Stop
        return [string]$t.State
    } catch {
        return 'not registered'
    }
}

function _Count-QueueJson([string]$QueueRoot, [string]$Folder) {
    $dir = Join-Path $QueueRoot $Folder
    if (-not (Test-Path -LiteralPath $dir)) { return 0 }
    return @((Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)).Count
}

function _Get-DashboardLockStatus([string]$QueueRoot) {
    $lockPath = Join-Path $QueueRoot '_dashboard.lock'
    if (-not (Test-Path -LiteralPath $lockPath)) {
        return @{ text = 'no lock'; alive = $false }
    }
    $pl = $null
    try { $pl = Get-Content -LiteralPath $lockPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $pl = $null }
    $pidVal = 0
    try { if ($pl -and $pl.pid) { $pidVal = [int]$pl.pid } } catch { $pidVal = 0 }
    if ($pidVal -le 0) { return @{ text = 'lock unreadable'; alive = $false } }
    $proc = Get-Process -Id $pidVal -ErrorAction SilentlyContinue
    if (-not $proc) { return @{ text = ('stale pid=' + $pidVal); alive = $false } }
    $isPw = ($proc.ProcessName -in @('powershell', 'pwsh'))
    if (-not $isPw) { return @{ text = ('pid=' + $pidVal + ' not powershell'); alive = $false } }
    return @{ text = ('alive pid=' + $pidVal); alive = $true }
}

function _Test-WatcherConsoleEvent {
    param([object]$Obj)
    if (-not $Obj) { return $false }
    $level = ''
    $code = ''
    try { $level = [string]$Obj.level } catch { }
    try { $code = [string]$Obj.code } catch { }
    if ($level -match '(?i)error|warning') { return $true }
    if ($code -match '(FAILED|STALE|ALERT|STALL|ACCEPTED|CONNECT_OK|WATCH_START$|ENQUEUE)') { return $true }
    return $false
}

function Write-QCDashboardWatcherLine {
    param([object]$Obj)
    if (-not (_Test-WatcherConsoleEvent -Obj $Obj)) { return $false }
    $code = [string]$Obj.code
    $msg = [string]$Obj.message
    $src = ''
    try {
        if ($Obj.data) {
            if ($Obj.data.sourceName) { $src = [string]$Obj.data.sourceName }
            elseif ($Obj.data.documentName) { $src = [string]$Obj.data.documentName }
            elseif ($Obj.data.jobType) { $src = [string]$Obj.data.jobType }
        }
    } catch { }
    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add(('[{0}]' -f (Get-Date -Format 'HH:mm:ss')))
    [void]$parts.Add('WATCH')
    if ($code) { [void]$parts.Add($code) }
    if ($src) { [void]$parts.Add($src) }
    if ($msg) { [void]$parts.Add($msg) }
    $color = 'Gray'
    $level = ''
    try { $level = [string]$Obj.level } catch { }
    if ($level -match '(?i)error' -or $code -match 'FAILED') { $color = 'Red' }
    elseif ($level -match '(?i)warning' -or $code -match 'STALE|ALERT') { $color = 'Yellow' }
    elseif ($code -match 'ACCEPTED|CONNECT_OK|WATCH_START') { $color = 'Cyan' }
    Write-Host ($parts -join ' ') -ForegroundColor $color
    return $true
}

$script:WatchJsonLog = @{
    hour = ''
    offset = [int64]0
    tail = ''
    skipExisting = $true
}

function _Get-WatcherJsonlPath {
    param([string]$LogDir, [string]$HourStamp)
    return (Join-Path $LogDir ("Watch-QCTrigger_${HourStamp}.jsonl"))
}

function Drain-QCDashboardWatcherJsonLogs {
    param([Parameter(Mandatory)][string]$LogDir)
    $hour = Get-QCLogHourStamp
    $state = $script:WatchJsonLog
    if ($state.hour -and [string]$state.hour -ne $hour) {
        $state.offset = [int64]0
        $state.tail = ''
        $state.skipExisting = $false
    }
    $state.hour = $hour
    $path = _Get-WatcherJsonlPath -LogDir $LogDir -HourStamp $hour
    if ($state.skipExisting) {
        if (Test-Path -LiteralPath $path) {
            try { $state.offset = [int64](Get-Item -LiteralPath $path).Length } catch { $state.offset = [int64]0 }
        }
        $state.skipExisting = $false
        return
    }
    $chunk = _QCRemoteWorkerReadLogChunkFromOffset -Path $path -StartPos ([int64]$state.offset)
    $state.offset = [int64]$chunk.NewPos
    $split = _QCRemoteWorkerProcessJsonlBuffer -Buffer (([string]$state.tail) + ([string]$chunk.Text))
    $state.tail = [string]$split.Remainder
    foreach ($line in ([string]$split.Complete -split '\r?\n')) {
        $t = ([string]$line).Trim()
        if (-not $t -or -not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            Write-QCDashboardWatcherLine -Obj $o | Out-Null
        } catch { }
    }
}

function Show-QCDashboardRecentWatcherLogs {
    param(
        [Parameter(Mandatory)][string]$LogDir,
        [int]$MaxLines = 20
    )
    $hour = Get-QCLogHourStamp
    $path = _Get-WatcherJsonlPath -LogDir $LogDir -HourStamp $hour
    if (-not (Test-Path -LiteralPath $path)) { return }
    $lines = @(Get-Content -LiteralPath $path -Tail ($MaxLines * 4) -ErrorAction SilentlyContinue)
    $shown = 0
    foreach ($line in $lines) {
        $t = ([string]$line).Trim()
        if (-not $t -or -not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            if (Write-QCDashboardWatcherLine -Obj $o) { $shown++ }
        } catch { }
        if ($shown -ge $MaxLines) { break }
    }
}

try { $host.UI.RawUI.WindowTitle = 'QC Pipeline Dashboard — logs' } catch { }

$cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath
$queueRoot = _Get-QueueRoot -Cfg $cfg
$logDir = _Get-ChildLogDir -QueueRoot $queueRoot
$hostName = [string]$env:COMPUTERNAME
$taskState = _Get-ScheduledTaskStateText -Name $ScheduledTaskName
$lockStatus = _Get-DashboardLockStatus -QueueRoot $queueRoot
$pendingN = _Count-QueueJson -QueueRoot $queueRoot -Folder 'pending'
$runningN = _Count-QueueJson -QueueRoot $queueRoot -Folder 'running'
$failedN = _Count-QueueJson -QueueRoot $queueRoot -Folder 'failed'

Initialize-QCRemoteWorkerHostLogView

Write-Host ("[{0}] QC pipeline log console host={1} queue={2}" -f (Get-Date -Format 'HH:mm:ss'), $hostName, $queueRoot) -ForegroundColor Cyan
Write-Host ("  scheduled task: {0} ({1})" -f $ScheduledTaskName, $taskState) -ForegroundColor Gray
Write-Host ("  dashboard lock: {0}" -f $lockStatus.text) -ForegroundColor $(if ($lockStatus.alive) { 'Green' } else { 'Yellow' })
Write-Host ("  queue: pending={0} running={1} failed={2}" -f $pendingN, $runningN, $failedN) -ForegroundColor Gray
Write-Host '  Read-only. Dashboard runs in the scheduled task (Session 0). Close this window anytime.' -ForegroundColor Gray

$seed = Sync-QCRemoteWorkerHostRunningView -QueueRoot $queueRoot -HostName $hostName -AllHosts -Quiet
$seedNames = @($seed.Jobs | ForEach-Object { if ($_.sourceName) { $_.sourceName } else { $_.jobId } })
if (@($seed.Jobs).Count -gt 0) {
    Write-Host ("  running now ({0}): {1}" -f @($seed.Jobs).Count, ($seedNames -join ', ')) -ForegroundColor Cyan
} else {
    Write-Host '  running now: none' -ForegroundColor DarkGray
}
Write-Host ''

if ($RecentLines -gt 0) {
    Write-Host '--- last watcher/worker JSONL lines ---' -ForegroundColor DarkGray
    Show-QCDashboardRecentWatcherLogs -LogDir $logDir -MaxLines ([Math]::Max(8, [int]($RecentLines / 2)))
    Show-QCRemoteWorkerHostRecentJsonLogs -LogDir $logDir -MaxLines $RecentLines
    Write-Host '--- live ---' -ForegroundColor DarkGray
}

$lastStatusAt = [DateTime]::MinValue

while ($true) {
    try { Drain-QCDashboardWatcherJsonLogs -LogDir $logDir } catch { }
    try { Drain-QCRemoteWorkerHostJsonLogs -LogDir $logDir } catch { }
    try { Sync-QCRemoteWorkerHostRunningView -QueueRoot $queueRoot -HostName $hostName -AllHosts | Out-Null } catch { }

    if (([DateTime]::UtcNow - $lastStatusAt).TotalSeconds -ge 10) {
        $lastStatusAt = [DateTime]::UtcNow
        $taskState = _Get-ScheduledTaskStateText -Name $ScheduledTaskName
        $lockStatus = _Get-DashboardLockStatus -QueueRoot $queueRoot
        $pendingN = _Count-QueueJson -QueueRoot $queueRoot -Folder 'pending'
        $busyTxt = Get-QCRemoteWorkerHostBusySummary -QueueRoot $queueRoot -HostName $hostName -AllHosts
        Write-Host ("[{0}] task={1} lock={2} pending={3} {4}" -f (Get-Date -Format 'HH:mm:ss'), $taskState, $lockStatus.text, $pendingN, $busyTxt) -ForegroundColor DarkGray
    }

    Start-Sleep -Milliseconds 400
}
