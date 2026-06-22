$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot 'diagnostics\Get-PWFolderStateCounts.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Error "Diagnostic script not found: $target"
    exit 1
}
& $target @args
exit $LASTEXITCODE
