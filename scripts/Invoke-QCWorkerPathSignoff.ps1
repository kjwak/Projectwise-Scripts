$target = Join-Path $PSScriptRoot 'diagnostics\Invoke-QCWorkerPathSignoff.ps1'
& $target @args
exit $LASTEXITCODE
