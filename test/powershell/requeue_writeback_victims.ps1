$ErrorActionPreference = 'Stop'

# Move the two STATUS_SET_GEN jobs that "succeeded" without actually uploading to
# PW back to the pending\ queue so the next worker run picks them up under the new
# code (which fails the job when writeback throws and persists pwUpload to disk).

$qroot = 'C:\QC_E2E_RealRun\queue'
$victims = @(
    'qc_statussetgen_6e3eff3e37add5c2'
    'qc_statussetgen_e86b5c1c81256b02'
)

foreach ($id in $victims) {
    $src = Join-Path $qroot ("succeeded\$id.json")
    $dst = Join-Path $qroot ("pending\$id.json")
    if (-not (Test-Path -LiteralPath $src)) {
        Write-Host "SKIP: $id not in succeeded\" -ForegroundColor Yellow
        continue
    }

    # Reset status + clear any stale result so the worker re-runs cleanly.
    $j = Get-Content -LiteralPath $src -Raw | ConvertFrom-Json
    $j.status = 'pending'
    if ($j.PSObject.Properties.Name -contains 'startedAtUtc') { $j.PSObject.Properties.Remove('startedAtUtc') }
    if ($j.PSObject.Properties.Name -contains 'result')        { $j.PSObject.Properties.Remove('result') }
    $j.updatedAtUtc = [DateTime]::UtcNow.ToString('o')
    $j | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $dst -Encoding UTF8
    Remove-Item -LiteralPath $src -Force
    Write-Host "REQUEUED: $id" -ForegroundColor Green
}

Write-Host ""
Write-Host "Pending count: " -NoNewline
(Get-ChildItem -LiteralPath (Join-Path $qroot 'pending') -Filter '*.json' -File).Count
Write-Host "Succeeded count: " -NoNewline
(Get-ChildItem -LiteralPath (Join-Path $qroot 'succeeded') -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
