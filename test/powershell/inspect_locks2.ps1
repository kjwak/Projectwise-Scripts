$qroot = 'C:\QC_E2E_RealRun\queue'
$locksDir = Join-Path $qroot 'locks'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Queue\QC.Queue.Json.psm1') -Force

Write-Host "=== Lock files - using the SAME _QCQJ-IsLockOwnerDead logic the acquire loop uses ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | ForEach-Object {
        $payload = $null
        try { $payload = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json } catch {}
        $ownerPid = if ($payload -and $payload.pid) { [int]$payload.pid } else { 0 }
        $isDead = & (Get-Module QC.Queue.Json) { param($p) _QCQJ-IsLockOwnerDead -LockPath $p } $_.FullName
        $age = ((Get-Date) - $_.LastWriteTime).TotalSeconds
        $alive = if ($ownerPid -gt 0) { Get-Process -Id $ownerPid -ErrorAction SilentlyContinue } else { $null }
        $aliveStr = if ($alive) { "ALIVE  $($alive.ProcessName)" } else { "DEAD" }
        Write-Host ("  {0,-46}  age={1,7:N1}s  pid={2,-6}  {3,-15}  IsLockOwnerDead={4}" -f $_.Name, $age, $ownerPid, $aliveStr, $isDead)
    }
