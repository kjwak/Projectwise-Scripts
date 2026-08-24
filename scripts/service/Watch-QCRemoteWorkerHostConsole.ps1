<#
.SYNOPSIS
Live console for the QC remote worker host scheduled task (read-only tail).

.DESCRIPTION
Does not start the supervisor. Tails queue\_logs\Run-QCProcessor_*.jsonl and prints
job activity from this host's RW* workers by default. Close the window anytime.

.EXAMPLE
.\scripts\service\Watch-QCRemoteWorkerHostConsole.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [string]$ScheduledTaskName = 'QC-RemoteWorkerHost',

    [Parameter(Mandatory = $false)]
    [int]$RecentLines = 40,

    [Parameter(Mandatory = $false)]
    [switch]$AllWorkers
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
    'Get-QCLogHourStamp'
) -Context 'remote worker log console'

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

try { $host.UI.RawUI.WindowTitle = 'QC Remote Worker Host — logs' } catch { }

$cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath
$rh = Get-QCRemoteWorkerHostSettings -Config $cfg
$queueRoot = _Get-QueueRoot -Cfg $cfg
$logDir = _Get-ChildLogDir -QueueRoot $queueRoot
$hostName = [string]$env:COMPUTERNAME
$types = if (@($rh.enabledJobTypes).Count -gt 0) { $rh.enabledJobTypes -join ',' } else { '(all)' }
$taskState = _Get-ScheduledTaskStateText -Name $ScheduledTaskName

Initialize-QCRemoteWorkerHostLogView -LocalWorkersOnly:(-not $AllWorkers.IsPresent)

Write-Host ("[{0}] QC remote worker log console host={1} queue={2}" -f (Get-Date -Format 'HH:mm:ss'), $hostName, $queueRoot) -ForegroundColor Cyan
Write-Host ("  scheduled task: {0} ({1})" -f $ScheduledTaskName, $taskState) -ForegroundColor Gray
Write-Host ("  filter: {0} types={1}" -f $(if ($AllWorkers.IsPresent) { 'all workers in JSONL' } else { 'RW* workers only (-AllWorkers for server lines too)' }), $types) -ForegroundColor Gray
Write-Host '  read-only tail — supervisor runs in the scheduled task. Close this window anytime.' -ForegroundColor Gray
Write-Host ''

if ($RecentLines -gt 0) {
    Write-Host ("--- last {0} matching lines ---" -f $RecentLines) -ForegroundColor DarkGray
    Show-QCRemoteWorkerHostRecentJsonLogs -LogDir $logDir -MaxLines $RecentLines
    Write-Host '--- live ---' -ForegroundColor DarkGray
}

$lastStatusAt = [DateTime]::MinValue

while ($true) {
    try { Drain-QCRemoteWorkerHostJsonLogs -LogDir $logDir } catch { }

    if (([DateTime]::UtcNow - $lastStatusAt).TotalSeconds -ge 10) {
        $lastStatusAt = [DateTime]::UtcNow
        $taskState = _Get-ScheduledTaskStateText -Name $ScheduledTaskName
        $busyTxt = Get-QCRemoteWorkerHostBusySummary
        Write-Host ("[{0}] task={1} {2}" -f (Get-Date -Format 'HH:mm:ss'), $taskState, $busyTxt) -ForegroundColor DarkGray
    }

    Start-Sleep -Milliseconds 400
}
