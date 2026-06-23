$qroot = 'C:\QC_E2E_RealRun\queue'
$locksDir = Join-Path $qroot 'locks'
foreach ($l in (Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue)) {
    $jobId = [System.IO.Path]::GetFileNameWithoutExtension($l.Name)
    if ($jobId -eq '_queue_write') { continue }
    $found = $null
    foreach ($s in 'pending','running','succeeded','failed') {
        $p = Join-Path (Join-Path $qroot $s) ($jobId + '.json')
        if (Test-Path -LiteralPath $p) { $found = $s; break }
    }
    Write-Host ("{0,-44}  state={1}" -f $l.Name, $found)
}
