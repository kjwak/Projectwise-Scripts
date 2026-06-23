$logDir = 'C:\QC_E2E_RealRun\queue\_logs'
$qroot = 'C:\QC_E2E_RealRun\queue'

Write-Host "=== Queue counts ===" -ForegroundColor Cyan
foreach ($s in 'pending','running','succeeded','failed') {
    $p = Join-Path $qroot $s
    if (Test-Path -LiteralPath $p) {
        $c = (Get-ChildItem -LiteralPath $p -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
        Write-Host ("  {0,-12} {1}" -f $s, $c)
    }
}
$lks = (Get-ChildItem -LiteralPath (Join-Path $qroot 'locks') -Filter '*.lock' -File -ErrorAction SilentlyContinue).Count
Write-Host ("  locks         {0}" -f $lks)

Write-Host "`n=== Most recent worker logs (top 4) ===" -ForegroundColor Cyan
$logs = Get-ChildItem -LiteralPath $logDir -Filter '*Run-QCProcessor*.out.log' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 4
foreach ($f in $logs) {
    Write-Host ("`n--- {0,-50}  size={1,8:N0}  lastWrite={2}" -f $f.Name, $f.Length, $f.LastWriteTime.ToString('HH:mm:ss')) -ForegroundColor Magenta
    $lines = Get-Content -LiteralPath $f.FullName -Tail 200 -ErrorAction SilentlyContinue
    foreach ($l in $lines) {
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        try {
            $j = $l | ConvertFrom-Json -ErrorAction Stop
            if ($j.code -match 'WORKER_(START|SELECTED|SUCCEEDED|FAILED|NO_JOB|LOCK_RACE|HANDLER|RELEASE|EXIT|BUDGET|LEASE)') {
                $jobId = if ($j.data -and $j.data.jobId) { $j.data.jobId } else { '' }
                Write-Host ("  {0}  {1,-22}  {2}  {3}" -f $j.ts.Substring(11,12), $j.code, $j.message, $jobId)
            }
        } catch { }
    }
}

Write-Host "`n=== Succeeded jobs (most recent 5) ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath (Join-Path $qroot 'succeeded') -Filter '*.json' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object {
        $j = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
        Write-Host ("  {0}  {1}  src={2}" -f $_.LastWriteTime.ToString('HH:mm:ss'), $j.id, $j.sourceFolder)
        if ($j.result) { Write-Host ("       result.code={0}  msg={1}" -f $j.result.code, $j.result.message) -ForegroundColor DarkGray }
    }

Write-Host "`n=== _StatusSet.pdf files (most recent 5) ===" -ForegroundColor Cyan
$ssRoot = 'C:\QC_E2E_RealRun\statussets'
if (Test-Path -LiteralPath $ssRoot) {
    Get-ChildItem -LiteralPath $ssRoot -Recurse -Filter '_StatusSet.pdf' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object {
            Write-Host ("  {0,12:N0}  {1}  {2}" -f $_.Length, $_.LastWriteTime.ToString('HH:mm:ss'), $_.FullName)
        }
}
