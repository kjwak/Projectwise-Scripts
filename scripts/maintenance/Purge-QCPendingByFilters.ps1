<#
.SYNOPSIS
Re-evaluates every job in queue\pending against the current appsettings.json
filters (whitelist/blacklist) and moves any disallowed job to queue\failed
with status='filtered'. Prints a summary.

.PARAMETER WhatIf
Show what would be purged without moving anything.

.PARAMETER AppSettingsPath
Path to appsettings.json. Defaults to <repo>\appsettings.json.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)

if (-not $AppSettingsPath) { $AppSettingsPath = Join-Path $repoRoot 'appsettings.json' }
if (-not (Test-Path -LiteralPath $AppSettingsPath)) { throw "appsettings.json not found: $AppSettingsPath" }

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Paths.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Queue\QC.Filters.psm1') -Force -DisableNameChecking | Out-Null

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$cfg = $cfgRes.Data.config

$queueRoot = $null
if ($cfg.queue) {
    if ($cfg.queue.rootDir) { $queueRoot = [string]$cfg.queue.rootDir }
    elseif ($cfg.queue.root) { $queueRoot = [string]$cfg.queue.root }
}
if (-not $queueRoot) { $queueRoot = Join-Path $repoRoot 'queue' }
if (-not (Test-Path -LiteralPath $queueRoot)) { throw "Queue root not found: $queueRoot" }
Write-Host ("Queue root: {0}" -f $queueRoot) -ForegroundColor DarkGray

$pendingDir = Join-Path $queueRoot 'pending'
$failedDir  = Join-Path $queueRoot 'failed'
if (-not (Test-Path -LiteralPath $pendingDir)) { Write-Host "No pending\ dir; nothing to purge." -ForegroundColor Yellow; return }
if (-not (Test-Path -LiteralPath $failedDir))  { New-Item -ItemType Directory -Path $failedDir -Force | Out-Null }

$files = @(Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File -ErrorAction SilentlyContinue)
Write-Host ("Scanning {0} pending job(s) against current filters..." -f $files.Count) -ForegroundColor Cyan

$purged = 0
$kept = 0
foreach ($f in $files) {
    $job = $null
    try { $job = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { $job = $null }
    if (-not $job) { continue }

    $candPath = $null
    if ($job.sourceFolder) { $candPath = [string]$job.sourceFolder }
    elseif ($job.sourcePath) { $candPath = [string]$job.sourcePath }
    elseif ($job.metadata -and $job.metadata.candidate -and $job.metadata.candidate.path) { $candPath = [string]$job.metadata.candidate.path }
    if (-not $candPath) { $kept++; continue }

    $allow = Test-QCPathAllowed -CandidatePath $candPath -Config $cfg
    if (-not $allow.IsSuccess) {
        Write-Host ("  ? {0,-40} {1}  (filter eval failed: {2})" -f $job.id, $candPath, $allow.Message) -ForegroundColor Yellow
        $kept++
        continue
    }

    if ([bool]$allow.Data.allowed) {
        $kept++
        continue
    }

    $reason = [string]$allow.Data.reason
    $rule = if ($allow.Data.matchedRule) { [string]$allow.Data.matchedRule } else { '' }
    Write-Host ("  PURGE {0}  type={1}  path='{2}'  reason={3} rule='{4}'" -f $job.id, $job.type, $candPath, $reason, $rule) -ForegroundColor Magenta

    if ($PSCmdlet.ShouldProcess($f.FullName, "Move pending->failed (filtered)")) {
        try {
            $job | Add-Member -MemberType NoteProperty -Name 'status' -Value 'failed' -Force
            $job | Add-Member -MemberType NoteProperty -Name 'filteredReason' -Value ("blacklist_repurge: $reason | $rule") -Force
            $job | Add-Member -MemberType NoteProperty -Name 'updatedAtUtc' -Value (Get-QCTimestamp) -Force
            $tmp = $f.FullName + '.tmp'
            $job | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tmp -Encoding utf8
            Move-Item -LiteralPath $tmp -Destination $f.FullName -Force
            $dst = Join-Path $failedDir $f.Name
            Move-Item -LiteralPath $f.FullName -Destination $dst -Force
            $purged++
        } catch {
            Write-Host ("    move failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    } else {
        $purged++
    }
}

Write-Host ""
Write-Host ("Done. purged={0}  kept={1}" -f $purged, $kept) -ForegroundColor Cyan
if ($WhatIfPreference) { Write-Host "(WhatIf mode - no files were moved)" -ForegroundColor DarkGray }
