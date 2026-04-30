<# Root shim for backwards compatibility.
Prefer running: scripts\run_prepend_qc.ps1
#>

$target = Join-Path $PSScriptRoot 'scripts\run_prepend_qc.ps1'
& $target @args
exit $LASTEXITCODE
