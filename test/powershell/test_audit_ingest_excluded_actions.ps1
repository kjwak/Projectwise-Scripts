$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.AuditPoller.psm1') -Force

$v = Get-AuditPollerLogicVersion
if ($v -ne '2026-08-24-folder-guid-cache-recon-only-v12') {
    throw "Expected audit poller logic v12, got: $v"
}

foreach ($code in @(1001, 1002, 1003, 1006, 1007, 1012, 1015, 1020)) {
    if (-not (Test-QCAuditIngestAllowedActionCode -ActionCode $code)) {
        throw "Action $code should be ingest-allowed"
    }
}
foreach ($code in @(1009, 1010, 1011, 1, 2, 2002)) {
    if (Test-QCAuditIngestAllowedActionCode -ActionCode $code) {
        throw "Action $code should not be ingest-allowed"
    }
}

Write-Host 'OK: only QCRelevantActions are ingest-allowed for audit_events' -ForegroundColor Green
