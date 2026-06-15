<#
.SYNOPSIS
Removes deprecated *-qc.pdf rows and references from QC_Pipeline database tables.

.DESCRIPTION
Legacy *-qc.pdf documents are replaced by lane PDFs (*-prod.pdf, *-chk.pdf, *-rev.pdf).
This script deletes registry rows for *-qc.pdf and clears qc_pdf_* columns that still
point at them. It does not delete ProjectWise documents.

Default is preview. Pass -ConfirmDeletes to apply.

.EXAMPLE
.\scripts\Remove-LegacyQcPdfDatabaseRows.ps1 -DryRun
.\scripts\Remove-LegacyQcPdfDatabaseRows.ps1 -ConfirmDeletes
.\scripts\Remove-LegacyQcPdfDatabaseRows.ps1 -ConfirmDeletes -FolderPathFilter 'CAFWY2200'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [string]$FolderPathFilter = '',
    [int]$BatchSize = 500
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

if ($DryRun.IsPresent -and $ConfirmDeletes.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmDeletes (apply changes), not both.'
}
if (-not $DryRun.IsPresent -and -not $ConfirmDeletes.IsPresent) {
    Write-Host 'Preview only: pass -ConfirmDeletes to apply, or -DryRun explicitly.' -ForegroundColor Yellow
    $DryRun = $true
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true.'
}

Initialize-QCDatabaseSchema -Config $config | Out-Null

function _RLQ-BuildFolderFilter {
    param([hashtable]$Params)
    if ([string]::IsNullOrWhiteSpace($FolderPathFilter)) { return '' }
    $Params['folderLike'] = ('%{0}%' -f $FolderPathFilter.Trim())
    return ' AND folder_path LIKE @folderLike'
}

function _RLQ-GetScalar {
    param([hashtable]$Config, [string]$Sql, [hashtable]$Params = @{})
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $Sql -Parameters $Params
    if (-not $res.IsSuccess) { throw $res.Message }
    if (-not $res.Data.table -or $res.Data.table.Rows.Count -eq 0) { return 0 }
    return [long]$res.Data.table.Rows[0][0]
}

function _RLQ-RunNonQuery {
    param([hashtable]$Config, [string]$Sql, [hashtable]$Params = @{}, [string]$Label = '')
    if ($DryRun) {
        Write-Host ("  [preview] {0}" -f $Label) -ForegroundColor DarkGray
        return 0
    }
    if (-not $PSCmdlet.ShouldProcess($Label, 'Execute')) { return 0 }
    $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $Sql -Parameters $Params
    if (-not $res.IsSuccess) { throw $res.Message }
    return [int]$res.Data.rowsAffected
}

$params = @{}
$folderClause = _RLQ-BuildFolderFilter -Params $params
$legacyNamePredicate = "LOWER(document_name) LIKE '%-qc.pdf'"

Write-Host "`n=== Legacy *-qc.pdf database cleanup ===" -ForegroundColor Cyan
if ($FolderPathFilter) { Write-Host ("Folder filter: {0}" -f $FolderPathFilter) -ForegroundColor DarkGray }

$counts = @{
    sheet_index_legacy_rows = _RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index
WHERE $legacyNamePredicate$folderClause
"@
    sheet_documents_legacy_rows = _RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_documents sd
WHERE LOWER(sd.document_name) LIKE '%-qc.pdf'
$(if ($folderClause) { " AND EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = sd.sheet_package_id$($folderClause -replace 'folder_path','sp.folder_path'))" } else { '' })
"@
    sheet_index_qc_link_rows = _RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index
WHERE qc_pdf_name IS NOT NULL AND LOWER(qc_pdf_name) LIKE '%-qc.pdf'$folderClause
"@
    sheet_packages_qc_link_rows = _RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_packages
WHERE qc_pdf_name IS NOT NULL AND LOWER(qc_pdf_name) LIKE '%-qc.pdf'$folderClause
"@
}

if (_RLQ-GetScalar -Config $config -Sql "SELECT CASE WHEN OBJECT_ID('sheet_package_qc_pdfs','U') IS NOT NULL THEN 1 ELSE 0 END" -Params @{}) {
    $counts['sheet_package_qc_pdfs_legacy_rows'] = _RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_package_qc_pdfs q
WHERE LOWER(q.document_name) LIKE '%-qc.pdf'
$(if ($folderClause) { " AND EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = q.sheet_package_id$($folderClause -replace 'folder_path','sp.folder_path'))" } else { '' })
"@
}

Write-Host "`nMatched rows:" -ForegroundColor Yellow
foreach ($k in $counts.Keys) { Write-Host ("  {0}: {1}" -f $k, $counts[$k]) }

if (($counts.Values | Measure-Object -Sum).Sum -eq 0) {
    Write-Host "`nNo legacy *-qc.pdf database rows matched." -ForegroundColor Green
    exit 0
}

if ($DryRun) {
    Write-Host "`nPreview complete. Re-run with -ConfirmDeletes to apply." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nApplying changes..." -ForegroundColor Cyan
$applied = @{
    sheet_index_qc_links_cleared = 0
    sheet_packages_qc_links_cleared = 0
    sheet_documents_deleted = 0
    sheet_package_qc_pdfs_deleted = 0
    sheet_index_deleted = 0
}

$applied['sheet_index_qc_links_cleared'] = _RLQ-RunNonQuery -Config $config -Label 'Clear sheet_index qc_pdf_* legacy links' -Params ($params + @{ batchSize = $BatchSize }) -Sql @"
UPDATE TOP (@batchSize) sheet_index
SET qc_pdf_guid = NULL, qc_pdf_name = NULL, last_updated_at = SYSDATETIMEOFFSET()
WHERE qc_pdf_name IS NOT NULL AND LOWER(qc_pdf_name) LIKE '%-qc.pdf'$folderClause
"@

$applied['sheet_packages_qc_links_cleared'] = _RLQ-RunNonQuery -Config $config -Label 'Clear sheet_packages qc_pdf_* legacy links' -Params ($params + @{ batchSize = $BatchSize }) -Sql @"
UPDATE TOP (@batchSize) sheet_packages
SET qc_pdf_guid = NULL, qc_pdf_name = NULL, last_updated_at = SYSDATETIMEOFFSET()
WHERE qc_pdf_name IS NOT NULL AND LOWER(qc_pdf_name) LIKE '%-qc.pdf'$folderClause
"@

while (_RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_documents sd
WHERE LOWER(sd.document_name) LIKE '%-qc.pdf'
$(if ($folderClause) { " AND EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = sd.sheet_package_id$($folderClause -replace 'folder_path','sp.folder_path'))" } else { '' })
"@ -gt 0) {
    $delDocSql = @"
DELETE TOP (@batchSize) sd FROM sheet_documents sd
WHERE LOWER(sd.document_name) LIKE '%-qc.pdf'
$(if ($folderClause) { " AND EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = sd.sheet_package_id$($folderClause -replace 'folder_path','sp.folder_path'))" } else { '' })
"@
    $n = _RLQ-RunNonQuery -Config $config -Label 'Delete sheet_documents legacy rows' -Params ($params + @{ batchSize = $BatchSize }) -Sql $delDocSql
    $applied['sheet_documents_deleted'] = [long]$applied['sheet_documents_deleted'] + $n
    if ($n -le 0) { break }
}

if ($counts.ContainsKey('sheet_package_qc_pdfs_legacy_rows') -and $counts['sheet_package_qc_pdfs_legacy_rows'] -gt 0) {
    while (_RLQ-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_package_qc_pdfs q
WHERE LOWER(q.document_name) LIKE '%-qc.pdf'
$(if ($folderClause) { " AND EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = q.sheet_package_id$($folderClause -replace 'folder_path','sp.folder_path'))" } else { '' })
"@ -gt 0) {
        $delLaneSql = @"
DELETE TOP (@batchSize) q FROM sheet_package_qc_pdfs q
WHERE LOWER(q.document_name) LIKE '%-qc.pdf'
$(if ($folderClause) { " AND EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = q.sheet_package_id$($folderClause -replace 'folder_path','sp.folder_path'))" } else { '' })
"@
        $n = _RLQ-RunNonQuery -Config $config -Label 'Delete sheet_package_qc_pdfs legacy rows' -Params ($params + @{ batchSize = $BatchSize }) -Sql $delLaneSql
        $applied['sheet_package_qc_pdfs_deleted'] = [long]$applied['sheet_package_qc_pdfs_deleted'] + $n
        if ($n -le 0) { break }
    }
}

while (_RLQ-GetScalar -Config $config -Params $params -Sql "SELECT COUNT_BIG(1) FROM sheet_index WHERE $legacyNamePredicate$folderClause" -gt 0) {
    $n = _RLQ-RunNonQuery -Config $config -Label 'Delete sheet_index legacy rows' -Params ($params + @{ batchSize = $BatchSize }) -Sql @"
DELETE TOP (@batchSize) FROM sheet_index
WHERE $legacyNamePredicate$folderClause
"@
    $applied['sheet_index_deleted'] = [long]$applied['sheet_index_deleted'] + $n
    if ($n -le 0) { break }
}

Write-Host "`nApplied:" -ForegroundColor Green
foreach ($k in ($applied.Keys | Sort-Object)) {
    if ($applied[$k]) { Write-Host ("  {0}: {1}" -f $k, $applied[$k]) }
}

Write-Host "`nLegacy *-qc.pdf database cleanup complete." -ForegroundColor Green
