<#
.SYNOPSIS
Forcibly terminates every PowerShell process that looks like a QC pipeline component
(dashboard, remote worker host, watcher, or worker) for this repository. Useful when stale dashboards
are running and Fortinet is throwing blocks because of the high concurrency.

.PARAMETER WhatIf
Show what would be killed without killing anything.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

$ErrorActionPreference = 'Stop'

$patterns = @(
    'Start-QCPipelineDashboard',
    'Start-QCRemoteWorkerHost',
    'Watch-QCTrigger',
    'Run-QCProcessor',
    'run_prepend_qc'
)
$rx = ($patterns -join '|')

$me = $PID
$procs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" |
    Where-Object { $_.CommandLine -and ($_.CommandLine -match $rx) -and ($_.ProcessId -ne $me) }

if (-not $procs -or @($procs).Count -eq 0) {
    Write-Host "No QC pipeline processes found." -ForegroundColor Green
    return
}

Write-Host ("Found {0} QC pipeline process(es):" -f @($procs).Count) -ForegroundColor Yellow
foreach ($p in $procs) {
    $cmd = if ($p.CommandLine.Length -gt 160) { $p.CommandLine.Substring(0,160) + '...' } else { $p.CommandLine }
    Write-Host ("  pid={0,-6} {1}" -f $p.ProcessId, $cmd)
}

if (-not $PSCmdlet.ShouldProcess("PIDs $((@($procs)|ForEach-Object{$_.ProcessId}) -join ',')", "Stop-Process -Force")) {
    Write-Host "(WhatIf - nothing killed)" -ForegroundColor DarkGray
    return
}

foreach ($p in $procs) {
    try {
        Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
        Write-Host ("  killed pid={0}" -f $p.ProcessId) -ForegroundColor Magenta
    } catch {
        Write-Host ("  FAILED to kill pid={0}: {1}" -f $p.ProcessId, $_.Exception.Message) -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Done. Re-run the dashboard once everything is clean." -ForegroundColor Cyan
