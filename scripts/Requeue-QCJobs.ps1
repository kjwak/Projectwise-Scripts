$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot 'maintenance\Requeue-QCJobs.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Error "Maintenance script not found: $target"
    exit 1
}
& $target @args
exit $LASTEXITCODE
