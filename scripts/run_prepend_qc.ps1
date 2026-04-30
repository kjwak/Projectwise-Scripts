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

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force

function _ToHashtable([object]$Value) {
  if ($null -eq $Value) { return $null }
  if ($Value -is [string]) { return $Value }
  if ($Value -is [System.ValueType]) { return $Value }
  if ($Value -is [System.Collections.IDictionary]) { return $Value }
  if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
    $out = @()
    foreach ($i in $Value) { $out += (_ToHashtable $i) }
    return $out
  }
  if ($Value.PSObject -and $Value.PSObject.Properties) {
    $h = @{}
    foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = (_ToHashtable $p.Value) }
    return $h
  }
  return $Value
}

function _Read-AppSettings([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    throw "appsettings.json not found: $Path"
  }
  $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
  $obj = $raw | ConvertFrom-Json -ErrorAction Stop
  return [hashtable](_ToHashtable $obj)
}

$watcher = Join-Path $PSScriptRoot 'Watch-QCTrigger.ps1'
$worker = Join-Path $PSScriptRoot 'Run-QCProcessor.ps1'

while ($true) {
  # 1) detect/enqueue
  & $watcher -AppSettingsPath $AppSettingsPath
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

  # 2) process queue until empty (bounded per poll)
  $cfg = _Read-AppSettings -Path $AppSettingsPath
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
