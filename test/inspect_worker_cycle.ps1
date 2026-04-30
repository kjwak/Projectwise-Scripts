$logDir = 'C:\QC_E2E_RealRun\queue\_logs'

# Find the latest log and show ALL events for a small time slice (no filter).
$latest = Get-ChildItem -LiteralPath $logDir -Filter '*Run-QCProcessor*.out.log' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $latest) { Write-Host "no logs"; exit 0 }
Write-Host "Reading $($latest.Name)" -ForegroundColor Cyan

# Pick the LAST 10 events. We need surrounding context, so take last ~60 lines.
$lines = Get-Content -LiteralPath $latest.FullName -Tail 200 -ErrorAction SilentlyContinue

# Pretty-print each JSON line as: ts code  msg  selected_id_if_present  (single line each)
foreach ($l in $lines) {
    if ([string]::IsNullOrWhiteSpace($l)) { continue }
    try {
        $j = $l | ConvertFrom-Json -ErrorAction Stop
        $ts = $j.ts
        $code = $j.code
        $level = $j.level
        $msg = $j.message
        $extra = ''
        if ($j.PSObject.Properties['data'] -and $j.data) {
            $d = $j.data
            $bits = @()
            foreach ($k in 'jobId','workerLabel','workerPid','jobType','errorCode','errorMessage','code') {
                if ($d.PSObject.Properties[$k] -and $d.$k) { $bits += "$k=$($d.$k)" }
            }
            if ($bits.Count) { $extra = ' [' + ($bits -join ', ') + ']' }
        }
        Write-Host ("{0}  {1,-5}  {2,-22}  {3}{4}" -f $ts.Substring(11,12), $level, $code, $msg, $extra)
    } catch {
        Write-Host "  (raw) $l" -ForegroundColor DarkGray
    }
}
