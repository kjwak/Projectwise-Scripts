$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Queue\QC.Queue.Json.psm1') -Force

$cfg = @{ queue = @{ root = 'C:\QC_E2E_RealRun\queue'; backend = 'json' } }

Write-Host "=== Before recovery ===" -ForegroundColor Cyan
foreach ($s in 'pending','running','succeeded','failed') {
    $p = Join-Path $cfg.queue.root $s
    $c = (Get-ChildItem -LiteralPath $p -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    Write-Host ("  {0,-12} {1}" -f $s, $c)
}
$lks = (Get-ChildItem -LiteralPath (Join-Path $cfg.queue.root 'locks') -Filter '*.lock' -File -ErrorAction SilentlyContinue).Count
Write-Host ("  locks         {0}" -f $lks)

Write-Host "`n=== Running Recover-QCStaleJobs ===" -ForegroundColor Cyan
$r = Recover-QCStaleJobs -Config $cfg
if (-not $r.IsSuccess) { Write-Host ("FAILED: " + $r.Message) -ForegroundColor Red; exit 1 }
$d = $r.Data
Write-Host ("  scanned (running):    {0}" -f $d.scanned)
Write-Host ("  recoveredToPending:   {0}" -f $d.recoveredToPending)
Write-Host ("  recoveredToFailed:    {0}" -f $d.recoveredToFailed)
Write-Host ("  recoveredOrphan:      {0}" -f $d.recoveredOrphan)
Write-Host ("  orphanLocksRemoved:   {0}" -f $d.orphanLocksRemoved)
foreach ($e in $d.details) { Write-Host ("    - {0}" -f ($e | ConvertTo-Json -Compress)) -ForegroundColor DarkGray }

Write-Host "`n=== After recovery ===" -ForegroundColor Cyan
foreach ($s in 'pending','running','succeeded','failed') {
    $p = Join-Path $cfg.queue.root $s
    $c = (Get-ChildItem -LiteralPath $p -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    Write-Host ("  {0,-12} {1}" -f $s, $c)
}
$lks = (Get-ChildItem -LiteralPath (Join-Path $cfg.queue.root 'locks') -Filter '*.lock' -File -ErrorAction SilentlyContinue).Count
Write-Host ("  locks         {0}" -f $lks)
