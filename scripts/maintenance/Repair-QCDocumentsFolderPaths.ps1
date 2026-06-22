<#
.SYNOPSIS
Normalize folder paths in QC_Pipeline tables to canonical documents\... form.

.DESCRIPTION
Uses Normalize-QCDocumentsFolderPath (same rules as queue jobs and DB writers).
Repairs base tables; views v_qc_cycle_aging and v_folder_activity update automatically.

.PARAMETER Table
Which tables to repair. "telemetry" = state/transition/notification/audit views sources.
"all" includes processing_jobs, sheet_index, document_activity.

.PARAMETER WhatIf
Report changes only; do not update.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$AppSettingsPath = '',
    [ValidateSet(
        'processing_jobs',
        'sheet_index',
        'audit_events',
        'document_state_history',
        'notification_log',
        'transition_events',
        'document_activity',
        'telemetry',
        'all'
    )]
    [string]$Table = 'telemetry',
    [switch]$WhatIf
)

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

$configRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $configRes.IsSuccess) { throw $configRes.Message }
$config = $configRes.Data.config

function _Repair-Column {
    param(
        [hashtable]$Config,
        [string]$TableName,
        [string]$KeyColumn,
        [string]$PathColumn
    )
    $sql = "SELECT $KeyColumn, $PathColumn FROM $TableName WHERE $PathColumn IS NOT NULL AND LTRIM(RTRIM($PathColumn)) <> ''"
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql
    if (-not $res.IsSuccess) {
        Write-Warning "Query failed for ${TableName}.${PathColumn}: $($res.Message)"
        return
    }
    $dt = $res.Data.table
    $changed = 0
    foreach ($dataRow in $dt.Rows) {
        $key = [string]$dataRow[$KeyColumn]
        $old = [string]$dataRow[$PathColumn]
        $normRes = Normalize-QCDocumentsFolderPath -Path $old
        if (-not $normRes.IsSuccess) { continue }
        $new = [string]$normRes.Data.path
        if ($new -eq $old) { continue }
        $changed++
        if ($WhatIf) {
            Write-Host "[WhatIf] $TableName $key : $old -> $new"
            continue
        }
        if ($PSCmdlet.ShouldProcess("$TableName.$PathColumn=$key", "Normalize path")) {
            $upd = "UPDATE $TableName SET $PathColumn = @p WHERE $KeyColumn = @k"
            [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql $upd -Parameters @{ p = $new; k = $key })
        }
    }
    Write-Host "$TableName.$PathColumn : $($dt.Rows.Count) rows scanned, $changed to update" -ForegroundColor Cyan
}

$targets = [System.Collections.Generic.List[hashtable]]::new()

function _AddTarget($t, $k, $p) {
    [void]$targets.Add(@{ t = $t; k = $k; p = $p })
}

switch ($Table) {
    'processing_jobs' {
        _AddTarget 'processing_jobs' 'job_id' 'source_folder'
        _AddTarget 'processing_jobs' 'job_id' 'source_path'
    }
    'sheet_index'         { _AddTarget 'sheet_index' 'document_guid' 'folder_path' }
    'audit_events'        { _AddTarget 'audit_events' 'id' 'resolved_folder' }
    'document_state_history' { _AddTarget 'document_state_history' 'id' 'folder_path' }
    'notification_log'    { _AddTarget 'notification_log' 'id' 'folder_path' }
    'transition_events'   { _AddTarget 'transition_events' 'id' 'folder_path' }
    'document_activity'   { _AddTarget 'document_activity' 'document_guid' 'folder_path' }
    'telemetry' {
        _AddTarget 'document_state_history' 'id' 'folder_path'
        _AddTarget 'notification_log' 'id' 'folder_path'
        _AddTarget 'transition_events' 'id' 'folder_path'
        _AddTarget 'audit_events' 'id' 'resolved_folder'
    }
    'all' {
        _AddTarget 'processing_jobs' 'job_id' 'source_folder'
        _AddTarget 'processing_jobs' 'job_id' 'source_path'
        _AddTarget 'sheet_index' 'document_guid' 'folder_path'
        _AddTarget 'audit_events' 'id' 'resolved_folder'
        _AddTarget 'document_state_history' 'id' 'folder_path'
        _AddTarget 'notification_log' 'id' 'folder_path'
        _AddTarget 'transition_events' 'id' 'folder_path'
        _AddTarget 'document_activity' 'document_guid' 'folder_path'
    }
}

Write-Host "Repair table scope: $Table ($($targets.Count) column targets)" -ForegroundColor Green
Write-Host "Views v_qc_cycle_aging / v_folder_activity are not updated directly (derived from base tables)." -ForegroundColor Yellow

foreach ($tgt in $targets) {
    _Repair-Column -Config $config -TableName $tgt.t -KeyColumn $tgt.k -PathColumn $tgt.p
}

Write-Host 'Done.'
