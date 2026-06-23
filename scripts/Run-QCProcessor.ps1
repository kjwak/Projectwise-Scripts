$target = Join-Path $PSScriptRoot 'service\Run-QCProcessor.ps1'
& $target @args
exit $LASTEXITCODE
