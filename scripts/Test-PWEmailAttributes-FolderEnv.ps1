$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot 'diagnostics\Test-PWEmailAttributes-FolderEnv.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Error "Diagnostic script not found: $target"
    exit 1
}
& $target @args
exit $LASTEXITCODE
