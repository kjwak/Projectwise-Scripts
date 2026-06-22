<#
.SYNOPSIS
Diagnostic snapshot of the QC queue: pending/running counts, per-running-job ages,
lock files with owner-PID liveness, and watcher-active sentinel state.

.PARAMETER QueueRoot
Optional explicit queue root. Defaults to reading queue.rootDir from appsettings.json.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$QueueRoot,
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
Import-Module (Join-Path $PSScriptRoot '..\..\modules\Core\Core.Runtime.psm1') -Force

if (-not $AppSettingsPath) { $AppSettingsPath = Join-Path $repoRoot 'appsettings.json' }

if (-not $QueueRoot) {
    if (-not (Test-Path -LiteralPath $AppSettingsPath)) { throw "appsettings.json not found: $AppSettingsPath" }
    $cfg = Get-Content -LiteralPath $AppSettingsPath -Raw | ConvertFrom-Json
    if ($cfg.queue -and $cfg.queue.rootDir) { $QueueRoot = [string]$cfg.queue.rootDir }
    elseif ($cfg.queue -and $cfg.queue.root) { $QueueRoot = [string]$cfg.queue.root }
    else { $QueueRoot = Join-Path $repoRoot 'queue' }
}

if (-not (Test-Path -LiteralPath $QueueRoot)) { throw "Queue root not found: $QueueRoot" }
Write-Host ("Queue root: {0}" -f $QueueRoot) -ForegroundColor Cyan

$now = (Get-Date).ToUniversalTime()

foreach ($s in @('pending','running','succeeded','failed')) {
    $dir = Join-Path $QueueRoot $s
    $count = 0
    if (Test-Path -LiteralPath $dir) {
        $count = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    }
    Write-Host ("  {0,-10} {1}" -f $s, $count)
}

Write-Host ""
Write-Host "== running jobs ==" -ForegroundColor Cyan
$runningDir = Join-Path $QueueRoot 'running'
if (Test-Path -LiteralPath $runningDir) {
    $rows = @()
    foreach ($f in Get-ChildItem -LiteralPath $runningDir -Filter '*.json' -File -ErrorAction SilentlyContinue) {
        $j = $null
        try { $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { $j = $null }
        $startedAtUtc = if ($j -and $j.startedAtUtc) { [string]$j.startedAtUtc } else { '' }
        $ageStartSecs = $null
        if ($startedAtUtc) {
            try { $ageStartSecs = [int]($now - ([DateTime]::Parse($startedAtUtc, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime())).TotalSeconds } catch { }
        }
        $ageDiskSecs = [int]($now - $f.LastWriteTimeUtc).TotalSeconds
        $rows += [pscustomobject]@{
            jobId         = if ($j) { [string]$j.id } else { [System.IO.Path]::GetFileNameWithoutExtension($f.Name) }
            type          = if ($j) { [string]$j.type } else { '' }
            ageStartSecs  = $ageStartSecs
            ageDiskSecs   = $ageDiskSecs
            sourceName    = if ($j) { [string]$j.sourceName } else { '' }
            sourceFolder  = if ($j) { [string]$j.sourceFolder } else { '' }
        }
    }
    if ($rows.Count -eq 0) { Write-Host '  (none)' -ForegroundColor DarkGray }
    else { $rows | Sort-Object -Property ageStartSecs -Descending | Format-Table -AutoSize | Out-String -Width 220 | Write-Host }
}

Write-Host "== locks (per-job, with owner PID liveness) ==" -ForegroundColor Cyan
$locksDir = Join-Path $QueueRoot 'locks'
if (Test-Path -LiteralPath $locksDir) {
    $rows = @()
    foreach ($f in Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue) {
        if ($f.Name -eq '_queue_write.lock') { continue }
        if ($f.Name -eq '_watcher_active.flag') { continue }
        $payload = $null
        try { $payload = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { $payload = $null }
        $ownerPid = if ($payload -and $payload.pid) { [int]$payload.pid } else { 0 }
        $alive = $false
        if ($ownerPid -gt 0) {
            $p = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
            if ($p) { $alive = $true }
        }
        $ageSecs = [int]($now - $f.LastWriteTimeUtc).TotalSeconds
        $rows += [pscustomobject]@{
            jobId        = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            ownerPid     = $ownerPid
            ownerAlive   = $alive
            ageSecs      = $ageSecs
            createdAtUtc = if ($payload) { [string]$payload.createdAtUtc } else { '' }
        }
    }
    if ($rows.Count -eq 0) { Write-Host '  (none)' -ForegroundColor DarkGray }
    else { $rows | Sort-Object -Property ageSecs -Descending | Format-Table -AutoSize | Out-String -Width 220 | Write-Host }
}

Write-Host "== _watcher_active.flag ==" -ForegroundColor Cyan
$flag = Join-Path $QueueRoot '_watcher_active.flag'
if (Test-Path -LiteralPath $flag) {
    $payload = Get-Content -LiteralPath $flag -Raw
    Write-Host $payload
} else {
    Write-Host '  (absent)' -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "== orphan analysis ==" -ForegroundColor Cyan
$running = @(Get-ChildItem -LiteralPath $runningDir -Filter '*.json' -File -ErrorAction SilentlyContinue)
$lockFiles = @(Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '_queue_write.lock' -and $_.Name -ne '_watcher_active.flag' })
$lockMap = @{}
foreach ($lf in $lockFiles) {
    $jid = [System.IO.Path]::GetFileNameWithoutExtension($lf.Name)
    $payload = $null
    try { $payload = Get-Content -LiteralPath $lf.FullName -Raw | ConvertFrom-Json } catch { }
    $ownerPid = 0
    if ($payload -and $payload.pid) { $ownerPid = [int]$payload.pid }
    $alive = $false
    if ($ownerPid -gt 0 -and (Get-Process -Id $ownerPid -ErrorAction SilentlyContinue)) { $alive = $true }
    $lockMap[$jid] = @{ ownerPid = $ownerPid; alive = $alive }
}
$orphans = @()
foreach ($f in $running) {
    $jid = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    if (-not $lockMap.ContainsKey($jid)) {
        $orphans += [pscustomobject]@{ jobId = $jid; reason = 'NO_LOCK_FILE' }
    } elseif (-not $lockMap[$jid].alive) {
        $orphans += [pscustomobject]@{ jobId = $jid; reason = "DEAD_PID($($lockMap[$jid].ownerPid))" }
    }
}
if ($orphans.Count -eq 0) { Write-Host '  No orphans detected (all running jobs have a live worker PID).' -ForegroundColor Green }
else {
    Write-Host ("  {0} orphaned running job(s):" -f $orphans.Count) -ForegroundColor Yellow
    $orphans | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
}
