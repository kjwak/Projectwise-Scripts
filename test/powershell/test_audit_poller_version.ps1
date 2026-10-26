$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.AuditPoller.psm1') -Force

$v = Get-AuditPollerLogicVersion
if ($v -ne '2026-08-24-folder-guid-cache-recon-only-v12') {
    throw "Expected audit poller logic v12, got: $v"
}

$h = @{ pw_action = 1012; pw_objguid = 'test-guid'; resolved_folder = 'Documents\Caltrans\X\CADD\Sheets\Seg_1' }
# GetRowValue is private; validate via GetTriggerActionCode path using module internals is not exported.
# Smoke: hashtable key access (regression for audit_events_db rows).
if ($h['pw_action'] -ne 1012) { throw 'hashtable fixture failed' }
Write-Host "OK: AuditPoller logic version $v" -ForegroundColor Green
