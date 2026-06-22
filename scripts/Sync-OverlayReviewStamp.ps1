$target = Join-Path $PSScriptRoot 'deployment\Sync-OverlayReviewStamp.ps1'
& $target @args
exit $LASTEXITCODE
