$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot 'maintenance\Remove-QCAuditEvents.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Error "Maintenance script not found: $target"
    exit 1
}
& $target @args
exit $LASTEXITCODE
