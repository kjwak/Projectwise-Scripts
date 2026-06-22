<#
.SYNOPSIS
Move jobs out of queue\succeeded (or queue\failed) back to queue\pending so they
will be picked up again by the worker pool. Useful after fixing a processor bug
that caused jobs to "succeed" without actually doing work.

.PARAMETER Type
Filter by job type (e.g., STATUS_SET_GEN, QC_PREPEND). If omitted, all types.

.PARAMETER From
Source state. One of 'succeeded' or 'failed'. Default 'succeeded'.

.PARAMETER WhatIf
Preview without moving anything.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$Type,
    [Parameter(Mandatory = $false)]
    [ValidateSet('succeeded','failed')]
    [string]$From = 'succeeded',
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
Import-Module (Join-Path $PSScriptRoot '..\..\modules\Core.Runtime.psm1') -Force
if (-not $AppSettingsPath) { $AppSettingsPath = Join-Path $repoRoot 'appsettings.json' }
if (-not (Test-Path -LiteralPath $AppSettingsPath)) { throw "appsettings.json not found: $AppSettingsPath" }

$cfg = Get-Content -LiteralPath $AppSettingsPath -Raw | ConvertFrom-Json
$queueRoot = $null
if ($cfg.queue -and $cfg.queue.rootDir) { $queueRoot = [string]$cfg.queue.rootDir }
elseif ($cfg.queue -and $cfg.queue.root) { $queueRoot = [string]$cfg.queue.root }
if (-not $queueRoot) { $queueRoot = Join-Path $repoRoot 'queue' }

$srcDir = Join-Path $queueRoot $From
$dstDir = Join-Path $queueRoot 'pending'
if (-not (Test-Path -LiteralPath $srcDir)) { Write-Host "Source dir not found: $srcDir"; return }
if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }

$files = @(Get-ChildItem -LiteralPath $srcDir -Filter '*.json' -File -ErrorAction SilentlyContinue)
$moved = 0
$skipped = 0
foreach ($f in $files) {
    $job = $null
    try { $job = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { $job = $null }
    if (-not $job) { $skipped++; continue }

    if ($Type -and ([string]$job.type) -ne $Type) { $skipped++; continue }

    $jobId = [string]$job.id
    $dst = Join-Path $dstDir ($jobId + '.json')

    Write-Host ("REQUEUE  {0,-40} type={1}  path='{2}'" -f $jobId, $job.type, $job.sourceFolder) -ForegroundColor Cyan

    if ($PSCmdlet.ShouldProcess($f.FullName, "Move $From -> pending")) {
        try {
            $job | Add-Member -MemberType NoteProperty -Name 'status' -Value 'pending' -Force
            $job | Add-Member -MemberType NoteProperty -Name 'startedAtUtc' -Value $null -Force
            $job | Add-Member -MemberType NoteProperty -Name 'updatedAtUtc' -Value (Get-QCTimestamp) -Force
            $tmp = $f.FullName + '.tmp'
            $job | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $tmp -Encoding utf8
            Move-Item -LiteralPath $tmp -Destination $f.FullName -Force
            Move-Item -LiteralPath $f.FullName -Destination $dst -Force
            $moved++
        } catch {
            Write-Host ("    move failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    } else {
        $moved++
    }
}

Write-Host ""
Write-Host ("Done. moved={0}  skipped={1}" -f $moved, $skipped) -ForegroundColor Green
if ($WhatIfPreference) { Write-Host "(WhatIf mode - no files were moved)" -ForegroundColor DarkGray }
