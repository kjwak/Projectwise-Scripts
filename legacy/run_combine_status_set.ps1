# run_combine_status_set.ps1
# Launcher for combine_status_set.ps1 across AZDOT Sheets folders.
# Runs continuously, monitoring for sheet updates and replacing them in each StatusSet.
# Discovers project\CADD\Sheets under Documents\AZDOT 2024 and Documents\AZDOT.
# Logs to C:\PW_QC_LOCAL\logs\
#
# IMPORTANT: Run from ProjectWise PowerShell (pwps) so pwps_dab is loaded.
#
# Run: .\run_combine_status_set.ps1
# One-shot: .\run_combine_status_set.ps1 -RunOnce
# Test PW columns: .\run_combine_status_set.ps1 -TestColumns

param(
  [Parameter(Mandatory = $false)]
  [switch] $RunOnce,

  [Parameter(Mandatory = $false)]
  [switch] $TestColumns
)

$scriptDir = $PSScriptRoot
$combineScript = Join-Path $scriptDir "combine_status_set.ps1"

# Pipe-delimited roots: powershell.exe -File does not bind string[] reliably; use -WatchUnderRootJoined.
$extra = @()
if ($RunOnce) { $extra += '-RunOnce' }
if ($TestColumns) { $extra += '-TestColumns' }

# Launch in MTA (pwps_dab requires it)
& powershell.exe -MTA -NoProfile -ExecutionPolicy Bypass -File $combineScript `
  -WatchUnderRootJoined 'Documents\AZDOT 2024|Documents\AZDOT' `
  -SheetsPathFromProject 'CADD\Sheets' `
  -WriteBackToPW `
  @extra
exit $LASTEXITCODE
