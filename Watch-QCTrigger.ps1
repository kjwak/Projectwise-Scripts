<# Root shim for backwards compatibility.
Prefer running: scripts\Watch-QCTrigger.ps1
#>

$target = Join-Path $PSScriptRoot 'scripts\Watch-QCTrigger.ps1'
& $target @args
exit $LASTEXITCODE

