$target = Join-Path $PSScriptRoot 'processing\Combine-StatusSet.ps1'
& $target @args
exit $LASTEXITCODE
