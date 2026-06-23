$target = Join-Path $PSScriptRoot 'service\Watch-QCTrigger.ps1'
& $target @args
exit $LASTEXITCODE
