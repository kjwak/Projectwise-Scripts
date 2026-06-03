$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

$v = Get-AuditPollerLogicVersion
if ($v -ne '2026-06-03-trigger-action-v2') {
    throw "Expected audit poller logic v2, got: $v"
}
Write-Host "OK: AuditPoller logic version $v" -ForegroundColor Green
