$target = Join-Path $PSScriptRoot 'service\Start-QCPipelineDashboard.ps1'
& $target @args
exit $LASTEXITCODE
