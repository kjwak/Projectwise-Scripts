$target = Join-Path $PSScriptRoot 'deployment\Promote-DevToMain.ps1'
& $target @args
exit $LASTEXITCODE
