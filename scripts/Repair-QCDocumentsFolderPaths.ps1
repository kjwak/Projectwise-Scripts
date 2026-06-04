<#
.SYNOPSIS
Normalize folder paths in QC_Pipeline tables to canonical documents\... form.

.DESCRIPTION
Uses Normalize-QCDocumentsFolderPath (same rules as new jobs and processing_jobs telemetry).
Safer than raw SQL when paths include pw: URIs or unusual casing.

.PARAMETER AppSettingsPath
Path to appsettings.json (for connection string).

.PARAMETER Table
processing_jobs (default), sheet_index, audit_events, or all.

.PARAMETER WhatIf
Report row counts only; do not update.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$AppSettingsPath = '',
    [ValidateSet('processing_jobs', 'sheet_index', 'audit_events', 'all')]
    [string]$Table = 'processing_jobs',
    [switch]$WhatIf
)

$repoRoot = Split-Path -Parent $PSScriptRoot
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

$targets = @()
switch ($Table) {
    'processing_jobs' { $targets += @{ t = 'processing_jobs'; k = 'job_id'; p = 'source_folder' }; $targets += @{ t = 'processing_jobs'; k = 'job_id'; p = 'source_path' } }
    'sheet_index'     { $targets += @{ t = 'sheet_index'; k = 'document_guid'; p = 'folder_path' } }
    'audit_events'    { $targets += @{ t = 'audit_events'; k = 'id'; p = 'resolved_folder' } }
    'all' {
        $targets += @{ t = 'processing_jobs'; k = 'job_id'; p = 'source_folder' }
        $targets += @{ t = 'processing_jobs'; k = 'job_id'; p = 'source_path' }
        $targets += @{ t = 'sheet_index'; k = 'document_guid'; p = 'folder_path' }
        $targets += @{ t = 'audit_events'; k = 'id'; p = 'resolved_folder' }
    }
}

foreach ($tgt in $targets) {
    _Repair-Column -Config $config -TableName $tgt.t -KeyColumn $tgt.k -PathColumn $tgt.p
}

Write-Host 'Done.'
