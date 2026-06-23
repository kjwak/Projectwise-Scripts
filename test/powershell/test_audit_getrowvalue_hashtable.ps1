# Regression: audit_events_db rows are hashtables with pw_action column names.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force

$table = New-Object System.Data.DataTable
[void]$table.Columns.Add('pw_action', [int])
[void]$table.Columns.Add('pw_objguid', [string])
$row = $table.NewRow()
$row['pw_action'] = 1012
$row['pw_objguid'] = '11111111-2222-3333-4444-555555555555'
[void]$table.Rows.Add($row)

$h = @(_QDB-ConvertDataTableToRowHashtables -Table $table)[0]
if ($h['pw_action'] -ne 1012) {
    throw "Converted hashtable missing pw_action (got $($h['pw_action']))"
}

Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.AuditPoller.psm1') -Force
if ((Get-AuditPollerLogicVersion) -ne '2026-06-03-hashtable-row-v3') {
    throw 'Expected audit poller logic v3'
}

Write-Host 'OK: DB hashtable row shape + poller v3.' -ForegroundColor Green
