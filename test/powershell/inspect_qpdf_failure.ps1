$logDir   = 'C:\QC_E2E_RealRun\queue\_logs'
$failedDir= 'C:\QC_E2E_RealRun\queue\failed'

Write-Host "=== Failed jobs (last 3) - lastResult / lastError ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath $failedDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending | Select-Object -First 3 | ForEach-Object {
    Write-Host "`n--- $($_.Name) ---" -ForegroundColor Magenta
    $j = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
    Write-Host "type=$($j.type)  attempts=$($j.attempts)  source=$($j.sourceFolder)"
    if ($j.lastError)  { Write-Host "lastError: $($j.lastError)" -ForegroundColor Yellow }
    if ($j.lastResult) {
        Write-Host "lastResult:" -ForegroundColor Yellow
        $j.lastResult | ConvertTo-Json -Depth 12
    }
}

Write-Host "`n=== Tail of latest worker logs - qpdf and STATUS_SET_QPDF lines ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath $logDir -Filter '*Run-QCProcessor*.out.log' -File -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending | Select-Object -First 3 | ForEach-Object {
    Write-Host "`n--- $($_.Name) ---" -ForegroundColor Magenta
    $lines = Get-Content -LiteralPath $_.FullName -Tail 400 -ErrorAction SilentlyContinue
    $lines | Where-Object { $_ -match 'STATUS_SET_QPDF|qpdf|MERGE|exitCode|stderr|stdout' } |
        Select-Object -Last 12 | ForEach-Object { Write-Host "  $_" }
}
