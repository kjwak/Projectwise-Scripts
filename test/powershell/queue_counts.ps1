$qroot = 'C:\QC_E2E_RealRun\queue'
foreach ($d in 'pending','running','succeeded','failed') {
    $p = Join-Path $qroot $d
    if (Test-Path $p) {
        $c = (Get-ChildItem -LiteralPath $p -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
        Write-Host ("{0,-12} {1}" -f $d, $c)
    }
}
