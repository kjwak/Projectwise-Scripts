$target = Join-Path $PSScriptRoot 'processing\Run-CombineStatusSet.ps1'
& $target @args
exit $LASTEXITCODE
