<# Root shim for backwards compatibility.
Prefer running: scripts\Run-QCProcessor.ps1
#>

$target = Join-Path $PSScriptRoot 'scripts\Run-QCProcessor.ps1'
& $target @args
exit $LASTEXITCODE
