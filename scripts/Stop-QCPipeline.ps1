$target = Join-Path $PSScriptRoot 'service\Stop-QCPipeline.ps1'
& $target @args
exit $LASTEXITCODE
