$logDir = 'C:\QC_E2E_RealRun\queue\_logs'
$failedDir = 'C:\QC_E2E_RealRun\queue\failed'

Write-Host "=== Recent failed jobs (full content of last 2) ===" -ForegroundColor Cyan
$ff = Get-ChildItem -LiteralPath $failedDir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 2
foreach ($f in $ff) {
    Write-Host ""
    Write-Host "--- $($f.Name) ---" -ForegroundColor Magenta
    $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
    Write-Host ("  type:    {0}" -f $j.type)
    Write-Host ("  attempts:{0}" -f $j.attempts)
    Write-Host ("  status:  {0}" -f $j.status)
    Write-Host ("  source:  {0}" -f $j.sourceFolder)
    if ($j.lastError) { Write-Host ("  lastError: {0}" -f $j.lastError) -ForegroundColor Yellow }
    if ($j.lastResult) {
        Write-Host "  lastResult:" -ForegroundColor Yellow
        Write-Host ($j.lastResult | ConvertTo-Json -Depth 8) -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "=== Tail of latest 2 Run-QCProcessor logs (last 80 lines, error/exception only) ===" -ForegroundColor Cyan
$logs = Get-ChildItem -LiteralPath $logDir -Filter '*Run-QCProcessor*' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 2
foreach ($f in $logs) {
    Write-Host ""
    Write-Host "--- $($f.Name) (size=$($f.Length)) ---" -ForegroundColor Magenta
    $tail = Get-Content -LiteralPath $f.FullName -Tail 200 -ErrorAction SilentlyContinue
    $tail | Where-Object { $_ -match 'WORKER_FAILED|WORKER_HANDLER_THREW|Error|Exception|throw|threw|fail|Forti|denied|UnauthorizedAccess|Access is denied|Could not load|invalid|parameter' } | Select-Object -Last 15 | ForEach-Object { Write-Host "  $_" }
}
