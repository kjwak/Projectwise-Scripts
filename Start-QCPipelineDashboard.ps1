<# Root entrypoint: live pipeline + dashboard.
Prefer running: scripts\Start-QCPipelineDashboard.ps1
#>

$target = Join-Path $PSScriptRoot 'scripts\Start-QCPipelineDashboard.ps1'
& $target @args
exit $LASTEXITCODE

