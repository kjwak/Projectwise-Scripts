# run_prepend_qc.ps1
# Launcher for prepend_qc_on_trigger with AZDOT 2024 + AZDOT Sheets folders (same roots as run_combine_status_set.ps1).
# Logs to C:\PW_QC_LOCAL\logs\ (activity + errors, daily rotation).
#
# IMPORTANT: Run from ProjectWise PowerShell (pwps) so pwps_dab is loaded.
#   Examples: Start menu > Bentley > ProjectWise PowerShell, or run from pwps prompt.
#
# Run: cd "C:\Users\jflint\Documents\ProjectWise Prepend"; .\run_prepend_qc.ps1
# One-shot: .\run_prepend_qc.ps1 -RunOnce
# Overlay: default per-sheet work folder under LocalRoot\work\<sheet>\ (split pages + MANIFEST). Override: -OverlaySheetWorkDir:$false -OverlayOldFromHistoryOnly:$true for qpdf temp --current-master only.

param(
  [Parameter(Mandatory = $false)]
  [switch] $RunOnce,

  # Add explicit watch folders (full pw:\...\path or relative folder paths).
  [Parameter(Mandatory = $false)]
  [string[]] $AddWatchFolderPaths,

  # Add more watch roots (same format as WatchUnderRootJoined entries, e.g. 'Documents\AZDOT 2024').
  [Parameter(Mandatory = $false)]
  [string[]] $AddWatchUnderRoot,

  # If set: scan each watched folder plus its immediate child folders.
  [Parameter(Mandatory = $false)]
  [switch] $OneLevelDeep,

  [Parameter(Mandatory = $false)]
  $OverlayOldFromHistoryOnly = $true,

  [Parameter(Mandatory = $false)]
  $OverlaySheetWorkDir = $true
)

$scriptDir = $PSScriptRoot
$triggerScript = Join-Path $scriptDir "prepend_qc_on_trigger.ps1"
$watchListPath = Join-Path $scriptDir "watchlist.json"
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

if ($OneLevelDeep) {
  $triggerParams['OneLevelDeep'] = $true
}

if ($allWatchFolders -and $allWatchFolders.Count -gt 0) {
  $triggerParams['ExtraWatchFolderPaths'] = @($allWatchFolders | Select-Object -Unique)
}

if ($rootsJoined -and $rootsJoined.Trim()) {
  $triggerParams['WatchUnderRootJoined'] = $rootsJoined
}
if ($RunOnce) { $triggerParams['RunOnce'] = $true }

& $triggerScript @triggerParams
exit $LASTEXITCODE
