$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot 'diagnostics\Test-PWEmailAttributes-EnvCount.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Error "Diagnostic script not found: $target"
    exit 1
}
& $target @args
exit $LASTEXITCODE
