# Regression: full ProjectWise audit action name map in PW.AuditPoller.psm1
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

if ((Get-PWAuditTrailActionName -ActionCode 1012) -ne 'DOCUMENT_STATE') {
    throw 'Expected DOCUMENT_STATE for 1012'
}

$checks = @{
    1023  = 'DOCUMENT_EXTRACT'
    7001  = 'WORKFLOW_CREATE'
    20001 = 'WORKAREA_CREATE'
    9104  = 'APPLICATION_VIEWER'
    6     = 'FOLDER_ACL_ASSIGN'
}
foreach ($code in $checks.Keys) {
    $name = Get-PWAuditTrailActionName -ActionCode $code
    if ($name -ne $checks[$code]) {
        throw "Code $code expected $($checks[$code]), got '$name'"
    }
}

if ((Get-PWAuditTrailActionName -ActionCode 99999) -ne 'UNKNOWN_99999') {
    throw 'Expected UNKNOWN_99999 for unmapped code'
}

Write-Host 'OK: audit action name map' -ForegroundColor Green
