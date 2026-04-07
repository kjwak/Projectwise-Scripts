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

  [Parameter(Mandatory = $false)]
  $OverlayOldFromHistoryOnly = $true,

  [Parameter(Mandatory = $false)]
  $OverlaySheetWorkDir = $true
)

$scriptDir = $PSScriptRoot
$triggerScript = Join-Path $scriptDir "prepend_qc_on_trigger.ps1"
# Hashtable splatting binds named parameters reliably (array form can fail in some hosts).
$triggerParams = @{
  WatchUnderRootJoined        = 'Documents\AZDOT 2024|Documents\AZDOT'
  SheetsPathFromProject       = 'CADD\Sheets'
  OverlayOldFromHistoryOnly   = $OverlayOldFromHistoryOnly
  OverlaySheetWorkDir         = $OverlaySheetWorkDir
}
if ($RunOnce) { $triggerParams['RunOnce'] = $true }

& $triggerScript @triggerParams
exit $LASTEXITCODE
