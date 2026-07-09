$qroot = 'C:\QC_E2E_RealRun\queue'
$locksDir = Join-Path $qroot 'locks'

Write-Host "=== Lock files in $locksDir ===" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $locksDir)) {
    Write-Host "  (locks dir does not exist)"
    exit 0
}
$locks = Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue
Write-Host ("  count: {0}" -f $locks.Count)
foreach ($l in ($locks | Sort-Object LastWriteTime -Descending)) {
    $body = (Get-Content -LiteralPath $l.FullName -Raw -ErrorAction SilentlyContinue)
    if ($null -eq $body) { $body = '<empty>' } else { $body = $body.Trim() }
    $age = ((Get-Date) - $l.LastWriteTime).TotalSeconds
    $pidAlive = '(no pid in body)'
    if ($body -match 'pid=(\d+)') {
        $p = [int]$matches[1]
        $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
        $pidAlive = if ($proc) { "ALIVE pid=$p name=$($proc.ProcessName)" } else { "DEAD pid=$p" }
    }
    Write-Host ("  {0,-50}  age={1,6:N1}s  body=[{2}]  {3}" -f $l.Name, $age, $body, $pidAlive)
}
