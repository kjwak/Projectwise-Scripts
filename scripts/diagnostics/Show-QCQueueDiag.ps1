<#
.SYNOPSIS
Diagnostic snapshot of the QC queue: pending/running counts, per-running-job ages,
lock files with owner-PID liveness, watcher-active sentinel state, and bounded
read-only log Warning/Error summary.

.PARAMETER QueueRoot
Optional explicit queue root. When omitted, uses the diagnostics location
resolver (env → local config → default UNC share). Does not default to
production appsettings queue.rootDir unless -UseAppSettingsQueueRoot is set.

.PARAMETER LogDir
Optional explicit log directory. When omitted, follows resolver precedence
(env → local config → {QueueRoot}\_logs).

.PARAMETER AppSettingsPath
Used only with -UseAppSettingsQueueRoot (escape hatch for local-queue debugging).

.PARAMETER UseAppSettingsQueueRoot
If set (and -QueueRoot omitted), resolve queue root from appsettings.json.

.PARAMETER NoLogs
Skip log directory inspection entirely.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$QueueRoot,

    [Parameter(Mandatory = $false)]
    [string]$LogDir,

    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [switch]$UseAppSettingsQueueRoot,

    [Parameter(Mandatory = $false)]
    [switch]$NoLogs
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Runtime.psm1'
) -Context 'Show-QCQueueDiag bootstrap'

. (Join-Path $PSScriptRoot 'Resolve-QCDebugLocations.ps1')

$resolveParams = @{
    RepoRoot = $repoRoot
}
if ($QueueRoot) { $resolveParams['QueueRoot'] = $QueueRoot }
if ($LogDir) { $resolveParams['LogDir'] = $LogDir }
if ($AppSettingsPath) { $resolveParams['AppSettingsPath'] = $AppSettingsPath }
if ($UseAppSettingsQueueRoot) { $resolveParams['UseAppSettingsQueueRoot'] = $true }

$locs = Resolve-QCDebugLocations @resolveParams
$QueueRoot = $locs.QueueRoot
$LogDir = $locs.LogDir

if (-not (Test-Path -LiteralPath $QueueRoot)) {
    throw ("Queue root not found or unreachable: {0} (source={1})" -f $QueueRoot, $locs.QueueRootSource)
}

Write-Host ("Queue root: {0}  (source={1})" -f $QueueRoot, $locs.QueueRootSource) -ForegroundColor Cyan
Write-Host ("Log dir:    {0}  (source={1})" -f $LogDir, $locs.LogDirSource) -ForegroundColor Cyan

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
if (-not (Test-Path -LiteralPath $locksDir)) {
    $altLocks = Join-Path $QueueRoot '_locks'
    if (Test-Path -LiteralPath $altLocks) { $locksDir = $altLocks }
}
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
} else {
    Write-Host '  (locks dir absent)' -ForegroundColor DarkGray
}

Write-Host "== _watcher_active.flag ==" -ForegroundColor Cyan
$flag = Join-Path $QueueRoot '_watcher_active.flag'
if (-not (Test-Path -LiteralPath $flag)) {
    $flagAlt = Join-Path $QueueRoot '_watcher\_watcher_active.flag'
    if (Test-Path -LiteralPath $flagAlt) { $flag = $flagAlt }
}
if (Test-Path -LiteralPath $flag) {
    $payload = Get-Content -LiteralPath $flag -Raw
    Write-Host $payload
} else {
    Write-Host '  (absent)' -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "== orphan analysis ==" -ForegroundColor Cyan
$running = @()
if (Test-Path -LiteralPath $runningDir) {
    $running = @(Get-ChildItem -LiteralPath $runningDir -Filter '*.json' -File -ErrorAction SilentlyContinue)
}
$lockFiles = @()
if (Test-Path -LiteralPath $locksDir) {
    $lockFiles = @(Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '_queue_write.lock' -and $_.Name -ne '_watcher_active.flag' })
}
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

if (-not $NoLogs) {
    Write-Host ""
    Write-Host "== logs (read-only, bounded) ==" -ForegroundColor Cyan
    if (-not (Test-Path -LiteralPath $LogDir)) {
        Write-Host ("  Log dir not found: {0} (source={1})" -f $LogDir, $locs.LogDirSource) -ForegroundColor Yellow
    } else {
        $allLogs = @(Get-ChildItem -LiteralPath $LogDir -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending)
        if ($allLogs.Count -eq 0) {
            Write-Host '  (empty)' -ForegroundColor DarkGray
        } else {
            Write-Host ("  Present: {0} file(s). Newest:" -f $allLogs.Count)
            $allLogs | Select-Object -First 5 | ForEach-Object {
                Write-Host ("    {0}  {1:u}" -f $_.Name, $_.LastWriteTimeUtc)
            }
        }

        $relevant = @(Get-ChildItem -LiteralPath $LogDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like 'Watch-QCTrigger_*.jsonl' -or $_.Name -like 'Run-QCProcessor_*.jsonl' } |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 5)

        if ($relevant.Count -eq 0) {
            Write-Host '  No Watch-QCTrigger_*.jsonl / Run-QCProcessor_*.jsonl files found.' -ForegroundColor DarkGray
        } else {
            Write-Host ("  Scanning newest {0} watcher/worker JSONL file(s), last 500 lines each (Warning/Error only):" -f $relevant.Count)
            $warnErr = @()
            foreach ($lf in $relevant) {
                $lines = @()
                try {
                    $lines = @(Get-Content -LiteralPath $lf.FullName -Tail 500 -ErrorAction Stop)
                } catch {
                    Write-Host ("    Failed to read {0}: {1}" -f $lf.Name, $_.Exception.Message) -ForegroundColor Yellow
                    continue
                }
                foreach ($line in $lines) {
                    if (-not $line) { continue }
                    $obj = $null
                    try { $obj = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }
                    $level = ''
                    if ($obj.PSObject.Properties['level']) { $level = [string]$obj.level }
                    elseif ($obj.PSObject.Properties['Level']) { $level = [string]$obj.Level }
                    elseif ($obj.PSObject.Properties['severity']) { $level = [string]$obj.severity }
                    if ($level -notmatch '^(?i)(warn|warning|error|err|fatal|critical)$') { continue }
                    $msg = ''
                    if ($obj.PSObject.Properties['message']) { $msg = [string]$obj.message }
                    elseif ($obj.PSObject.Properties['msg']) { $msg = [string]$obj.msg }
                    elseif ($obj.PSObject.Properties['Message']) { $msg = [string]$obj.Message }
                    $ts = ''
                    if ($obj.PSObject.Properties['ts']) { $ts = [string]$obj.ts }
                    elseif ($obj.PSObject.Properties['timestamp']) { $ts = [string]$obj.timestamp }
                    elseif ($obj.PSObject.Properties['time']) { $ts = [string]$obj.time }
                    $warnErr += [pscustomobject]@{
                        file    = $lf.Name
                        ts      = $ts
                        level   = $level
                        message = if ($msg.Length -gt 160) { $msg.Substring(0, 157) + '...' } else { $msg }
                    }
                }
            }
            if ($warnErr.Count -eq 0) {
                Write-Host '  No Warning/Error entries in scanned window.' -ForegroundColor Green
            } else {
                Write-Host ("  {0} Warning/Error entr(ies):" -f $warnErr.Count) -ForegroundColor Yellow
                $warnErr | Select-Object -Last 40 | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
            }
        }
    }
} else {
    Write-Host ""
    Write-Host "== logs ==" -ForegroundColor Cyan
    Write-Host '  Skipped (-NoLogs).' -ForegroundColor DarkGray
}
