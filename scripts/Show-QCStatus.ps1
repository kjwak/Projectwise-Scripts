<#
.SYNOPSIS
Terminal-friendly QC status snapshot.

.DESCRIPTION
Reads appsettings.json and prints queue stats + recent jobs (no mutations).
This is a first building block for a richer dashboard later.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json'),

    [Parameter(Mandatory = $false)]
    [int]$Recent = 15
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$stats = Get-QCQueueStats -Config $config
if (-not $stats.IsSuccess) { throw $stats.Message }

$root = [string]$stats.Data.root
$states = $stats.Data.states
$locks = $stats.Data.locks

Write-Host ""
Write-Host "QC Queue Root: $root" -ForegroundColor Cyan
Write-Host ("Pending:   {0,6}" -f [int]$states.pending)
Write-Host ("Running:   {0,6}" -f [int]$states.running)
Write-Host ("Succeeded: {0,6}" -f [int]$states.succeeded)
Write-Host ("Failed:    {0,6}" -f [int]$states.failed)
Write-Host ("Locks:     {0,6}" -f [int]$locks.count)

$recentRes = Get-QCRecentJobs -Config $config -Limit $Recent
if (-not $recentRes.IsSuccess) { throw $recentRes.Message }

Write-Host ""
Write-Host "Recent jobs (newest first):" -ForegroundColor Cyan
foreach ($entry in @($recentRes.Data.jobs)) {
    # Get-QCRecentJobs returns entries shaped like: @{ state; job; lastWriteTimeUtc }.
    # For backward compatibility, also tolerate a flat job object.
    $job = $null
    try { if ($entry -and $entry.job) { $job = $entry.job } } catch { $job = $null }
    if (-not $job) { $job = $entry }

    $id = ''
    $typ = ''
    $state = ''
    $ts = ''
    try { if ($job -and $job.id) { $id = [string]$job.id } } catch { }
    try { if ($job -and $job.type) { $typ = [string]$job.type } } catch { }
    try { if ($entry -and $entry.state) { $state = [string]$entry.state } elseif ($job -and $job.status) { $state = [string]$job.status } } catch { }
    try { if ($entry -and $entry.lastWriteTimeUtc) { $ts = [string]$entry.lastWriteTimeUtc } elseif ($job -and $job.updatedAtUtc) { $ts = [string]$job.updatedAtUtc } } catch { }

    Write-Host ("{0}  {1,-9}  {2,-14}  {3}" -f $ts, $state, $typ, $id)
}
Write-Host ""

