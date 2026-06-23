$qroot = 'C:\QC_E2E_RealRun\queue'
Get-ChildItem -LiteralPath (Join-Path $qroot 'succeeded') -Filter '*.json' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 2 | ForEach-Object {
        Write-Host "`n=== $($_.Name) ===" -ForegroundColor Magenta
        $j = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
        $j | ConvertTo-Json -Depth 12
    }
