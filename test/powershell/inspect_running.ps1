$qroot = 'C:\QC_E2E_RealRun\queue'
$logDir = Join-Path $qroot '_logs'

Write-Host "=== Running jobs (full json + lock state + age) ===" -ForegroundColor Cyan
$running = Get-ChildItem -LiteralPath (Join-Path $qroot 'running') -Filter '*.json' -File -ErrorAction SilentlyContinue
foreach ($f in $running) {
    Write-Host "`n--- $($f.Name) ---" -ForegroundColor Magenta
    $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
    Write-Host ("type:        {0}" -f $j.type)
    Write-Host ("source:      {0}" -f $j.sourceFolder)
    Write-Host ("attempts:    {0}" -f $j.attempts)
    Write-Host ("status:      {0}" -f $j.status)
    if ($j.PSObject.Properties['runningSince']) { Write-Host ("runningSince:{0}" -f $j.runningSince) }
    if ($j.PSObject.Properties['leaseExpiresAt']){ Write-Host ("leaseExp:    {0}" -f $j.leaseExpiresAt) }
    if ($j.PSObject.Properties['workerLabel'])  { Write-Host ("workerLabel: {0}" -f $j.workerLabel) }
    if ($j.PSObject.Properties['workerPid'])    { Write-Host ("workerPid:   {0}" -f $j.workerPid) }
    Write-Host ("ageMin:      {0:N1}" -f ((Get-Date).ToUniversalTime() - $f.LastWriteTime.ToUniversalTime()).TotalMinutes)
    $lock = (Join-Path $qroot 'running') + "\$($j.id).lock"
    if (Test-Path -LiteralPath $lock) {
        $body = (Get-Content -LiteralPath $lock -Raw -ErrorAction SilentlyContinue)
        Write-Host ("lock:        {0}" -f $body)
        if ($body -match 'pid=(\d+)') {
            $p = [int]$matches[1]
            $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
            Write-Host ("lock pid alive: {0}" -f ([bool]$proc))
        }
    } else {
        Write-Host "lock:        (none)" -ForegroundColor Yellow
    }
}

Write-Host "`n=== _StatusSet.pdf files in statussets dir ===" -ForegroundColor Cyan
$ssRoot = 'C:\QC_E2E_RealRun\statussets'
if (Test-Path -LiteralPath $ssRoot) {
    Get-ChildItem -LiteralPath $ssRoot -Recurse -Filter '_StatusSet.pdf' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 10 | ForEach-Object {
            Write-Host ("  {0,12:N0}  {1}  {2}" -f $_.Length, $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'), $_.FullName)
        }
}

Write-Host "`n=== Latest worker log lines mentioning STATUS_SET_GEN_DONE / Update-PW / New-PW / writeBack ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath $logDir -Filter '*Run-QCProcessor*.out.log' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 4 | ForEach-Object {
        Write-Host "`n--- $($_.Name) ---" -ForegroundColor Magenta
        Get-Content -LiteralPath $_.FullName -Tail 800 -ErrorAction SilentlyContinue |
            Where-Object { $_ -match 'STATUS_SET|writeBack|Update-PW|New-PW|MERGE_OK|UPLOAD|UPDATE_PW|NEW_PW|WORKER_(SELECTED|DONE|FAILED|RELEASE|FINISHED)|JOB_(LEASE|RELEASE)' } |
            Select-Object -Last 30 | ForEach-Object { Write-Host "  $_" }
    }
