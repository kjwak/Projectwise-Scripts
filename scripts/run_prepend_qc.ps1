$target = Join-Path $PSScriptRoot 'service\run_prepend_qc.ps1'
& $target @args
exit $LASTEXITCODE
