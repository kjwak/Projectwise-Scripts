<#
.SYNOPSIS
Remove duplicate queue job JSON files (same job id in pending, running, succeeded, etc.).
.DESCRIPTION
Keeps the copy in the most advanced state: succeeded > running > pending > failed.
Run after AV blocked queue file deletes, or when the dashboard shows stale jobs in multiple columns.
.PARAMETER WhatIf
Preview removals without deleting files.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath
)

$ErrorActionPreference = 'Stop'
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Queue\QC.Queue.Json.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
) -Context 'Repair-QCQueueDuplicates bootstrap'

if (-not $AppSettingsPath) { $AppSettingsPath = Join-Path $repoRoot 'appsettings.json' }
$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$cfg = [hashtable]$cfgRes.Data.config
$res = Repair-QCQueueDuplicateJobs -Config $cfg
if (-not $res.IsSuccess) { throw ($res.Code + ': ' + $res.Message) }

$repaired = @($res.Data.repaired)
$failed = @($res.Data.removeFailed)
Write-Host ("Removed {0} duplicate file(s)." -f $repaired.Count)
foreach ($r in $repaired) {
    Write-Host ("  {0}: removed {1}, kept {2}" -f $r.jobId, $r.removed, $r.kept)
}
if ($failed.Count -gt 0) {
    Write-Host ("Failed to remove {0} file(s) (AV may still hold locks):" -f $failed.Count) -ForegroundColor Yellow
    foreach ($f in $failed) {
        Write-Host ("  {0}: {1}" -f $f.jobId, $f.path)
    }
}
