# run_prepend_qc.ps1
# Launcher for prepend_qc_on_trigger with AZDOT 2024 Sheets folders.
# Logs to C:\PW_QC_LOCAL\logs\ (activity + errors, daily rotation).
#
# IMPORTANT: Run from ProjectWise PowerShell (pwps) so pwps_dab is loaded.
#   Examples: Start menu > Bentley > ProjectWise PowerShell, or run from pwps prompt.
#
# Run: cd "C:\Users\jflint\Documents\ProjectWise Prepend"; .\run_prepend_qc.ps1
# One-shot: .\run_prepend_qc.ps1 -RunOnce

param(
  [Parameter(Mandatory = $false)]
  [switch] $RunOnce
)

$scriptDir = $PSScriptRoot
$triggerScript = Join-Path $scriptDir "prepend_qc_on_trigger.ps1"
$params = @(
  '-WatchUnderRoot', 'Documents\AZDOT 2024',
  '-SheetsPathFromProject', 'CADD\Sheets'
)
if ($RunOnce) { $params += '-RunOnce' }

& $triggerScript @params
exit $LASTEXITCODE
