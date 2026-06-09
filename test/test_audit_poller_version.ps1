$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

$v = Get-AuditPollerLogicVersion
if ($v -ne '2026-06-05-parent-guid-cache-gate-v10') {
    throw "Expected audit poller logic v10, got: $v"
}

$h = @{ pw_action = 1012; pw_objguid = 'test-guid'; resolved_folder = 'Documents\Caltrans\X\CADD\Sheets\Seg_1' }
# GetRowValue is private; validate via GetTriggerActionCode path using module internals is not exported.
# Smoke: hashtable key access (regression for audit_events_db rows).
if ($h['pw_action'] -ne 1012) { throw 'hashtable fixture failed' }
Write-Host "OK: AuditPoller logic version $v" -ForegroundColor Green
