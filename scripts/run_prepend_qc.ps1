<# 
run_prepend_qc.ps1

Launcher for the new modular pipeline:
  Watch-QCTrigger.ps1 (detect/enqueue) + Run-QCProcessor.ps1 (dequeue/process).

This replaces the legacy "prepend_qc_on_trigger.ps1" monolith while preserving
an opt-in compatibility path via -Legacy.
#>

param(
  [Parameter(Mandatory = $false)]
  [switch] $RunOnce,

  [Parameter(Mandatory = $false)]
  [string] $AppSettingsPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json'),

  # Run the terminal dashboard (recommended; streams live scan/queue/proc status).
  # Use -NoDashboard to run watcher/worker in the foreground (legacy "no UI" behavior).
  [Parameter(Mandatory = $false)]
  [switch] $NoDashboard,

  [Parameter(Mandatory = $false)]
  [int] $PollSeconds = 0,

  [Parameter(Mandatory = $false)]
  [int] $MaxJobsPerPoll = 50,

  # Use the legacy monolith entrypoint.
  [Parameter(Mandatory = $false)]
  [switch] $Legacy,

  # Legacy params (only used with -Legacy).
  [Parameter(Mandatory = $false)]
  [string[]] $AddWatchFolderPaths,
  [Parameter(Mandatory = $false)]
  [string[]] $AddWatchUnderRoot,
  [Parameter(Mandatory = $false)]
  [switch] $OneLevelDeep,
  [Parameter(Mandatory = $false)]
  $OverlayOldFromHistoryOnly = $true,
  [Parameter(Mandatory = $false)]
  $OverlaySheetWorkDir = $true
)

$repoRoot = Split-Path -Parent $PSScriptRoot

if ($Legacy) {
  $triggerScript = Join-Path $repoRoot "legacy\\prepend_qc_on_trigger.ps1"
  $watchListPath = Join-Path $repoRoot "legacy\\watchlist.json"
  # Hashtable splatting binds named parameters reliably (array form can fail in some hosts).
  $defaultRoots = @()
  $allRoots = @($defaultRoots + @($AddWatchUnderRoot) | ForEach-Object { ($_ -as [string]).Trim() } | Where-Object { $_ })
  $rootsJoined = ($allRoots | Select-Object -Unique) -join '|'
  $allWatchFolders = @(@($AddWatchFolderPaths) | ForEach-Object { ($_ -as [string]).Trim() } | Where-Object { $_ })
  $triggerParams = @{
    WatchListPath              = $watchListPath
    SheetsPathFromProject       = 'CADD\Sheets'
    OverlayOldFromHistoryOnly   = $OverlayOldFromHistoryOnly
    OverlaySheetWorkDir         = $OverlaySheetWorkDir
  }
  if ($OneLevelDeep) { $triggerParams['OneLevelDeep'] = $true }
  if ($allWatchFolders -and $allWatchFolders.Count -gt 0) { $triggerParams['ExtraWatchFolderPaths'] = @($allWatchFolders | Select-Object -Unique) }
  if ($rootsJoined -and $rootsJoined.Trim()) { $triggerParams['WatchUnderRootJoined'] = $rootsJoined }
  if ($RunOnce) { $triggerParams['RunOnce'] = $true }
  & $triggerScript @triggerParams
  exit $LASTEXITCODE
}

if (-not $NoDashboard) {
  $dash = Join-Path $repoRoot 'scripts\Start-QCPipelineDashboard.ps1'
  $dashArgs = @('-AppSettingsPath', $AppSettingsPath, '-PollSeconds', $PollSeconds, '-MaxJobsPerPoll', $MaxJobsPerPoll)
  if ($RunOnce) { $dashArgs += '-PollSeconds'; $dashArgs += 0 }
  & $dash @dashArgs
  exit $LASTEXITCODE
}

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Queue\QC.Queue.Json.psm1') -Force

$watcher = Join-Path $PSScriptRoot 'Watch-QCTrigger.ps1'
$worker = Join-Path $PSScriptRoot 'Run-QCProcessor.ps1'

$cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath
try {
  $startupRes = Invoke-QCQueueStartupCheck -Config $cfg -ClearWatcherActive
  $startupData = if ($startupRes.IsSuccess) { $startupRes.Data } else { @{} }
  $qStates = $null
  if ($startupData.queueStats -and $startupData.queueStats.states) { $qStates = $startupData.queueStats.states }
  $pend = 0; $run = 0
  if ($qStates) {
    try { $pend = [int]$qStates.pending } catch { }
    try { $run = [int]$qStates.running } catch { }
  }
  $rec = $startupData.recovery
  $requeued = 0; $failedRec = 0; $orphans = 0
  if ($rec) {
    try { $requeued = [int]$rec.recoveredToPending } catch { }
    try { $failedRec = [int]$rec.recoveredToFailed } catch { }
    try { $orphans = [int]$rec.recoveredOrphan } catch { }
  }
  Write-Host ("Queue startup: pending={0} running={1} | recovery requeued={2} failed={3} orphans={4}" -f $pend, $run, $requeued, $failedRec, $orphans) -ForegroundColor Cyan
  if ($startupData.errors -and @($startupData.errors).Count -gt 0) {
    Write-Warning ("Queue startup partial errors: {0}" -f (($startupData.errors | ForEach-Object { [string]$_ }) -join '; '))
  }
} catch {
  Write-Warning ("Queue startup check failed: {0}" -f $_.Exception.Message)
}

while ($true) {
  # 1) detect/enqueue
  & $watcher -AppSettingsPath $AppSettingsPath
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

  # 2) process queue until empty (bounded per poll)
  $processed = 0
  while ($processed -lt $MaxJobsPerPoll) {
    $stats = Get-QCQueueStats -Config $cfg
    if (-not $stats.IsSuccess) { throw $stats.Message }
    $pending = 0
    try { $pending = [int]$stats.Data.states.pending } catch { $pending = 0 }
    if ($pending -le 0) { break }

    & $worker -AppSettingsPath $AppSettingsPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    $processed++
  }

  if ($RunOnce) { break }
  if ($PollSeconds -gt 0) { Start-Sleep -Seconds $PollSeconds }
}

exit 0
