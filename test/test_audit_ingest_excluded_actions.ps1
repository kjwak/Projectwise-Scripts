$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

$v = Get-AuditPollerLogicVersion
if ($v -ne '2026-06-04-ingest-exclude-checkout-v6') {
    throw "Expected audit poller logic v6, got: $v"
}

foreach ($code in @(1009, 1010, 1011)) {
    if (-not (Test-QCAuditIngestExcludedActionCode -ActionCode $code)) {
        throw "Action $code should be ingest-excluded"
    }
}
foreach ($code in @(1001, 1002, 1003, 1006, 1007, 1012, 1015, 1020)) {
    if (Test-QCAuditIngestExcludedActionCode -ActionCode $code) {
        throw "Action $code should not be ingest-excluded"
    }
}

Write-Host 'OK: checkout actions excluded from audit_events ingest' -ForegroundColor Green
