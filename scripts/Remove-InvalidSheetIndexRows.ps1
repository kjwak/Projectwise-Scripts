<#
.SYNOPSIS
Removes sheet_index rows whose document_guid is too short to be a ProjectWise GUID.

.DESCRIPTION
Early sheet_index imports sometimes stored document_number (or another numeric id) in
document_guid, producing 3-digit keys like "258" alongside the real PW GUID for the same
document_name + folder_path. Those rows are useless for PW lookups and appear as duplicates
in v_sheet_status (view over sheet_index).

Deletes from dbo.sheet_index only. v_sheet_status has no separate storage.

Default is preview (-DryRun). Pass -ConfirmDeletes to apply.

.EXAMPLE
.\scripts\Remove-InvalidSheetIndexRows.ps1 -DryRun
.\scripts\Remove-InvalidSheetIndexRows.ps1 -ConfirmDeletes -FolderPathFilter 'Documents\AZDOT\AZFWY2302-018'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [string]$FolderPathFilter = '',
    [int]$MinGuidLength = 5,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

if ($DryRun.IsPresent -and $ConfirmDeletes.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmDeletes (apply deletes), not both.'
}

$doDelete = $ConfirmDeletes.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmDeletes.IsPresent) {
    Write-Host 'Preview only: pass -ConfirmDeletes to delete rows, or -DryRun explicitly.' -ForegroundColor Yellow
    $DryRun = $true
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true.'
}

Initialize-QCDatabaseSchema -Config $config | Out-Null

$folderClause = ''
$params = @{ minLen = $MinGuidLength }
if (-not [string]::IsNullOrWhiteSpace($FolderPathFilter)) {
    $folderClause = ' AND folder_path LIKE @folderLike'
    $params['folderLike'] = ('%{0}%' -f $FolderPathFilter.Trim())
}

$previewSql = @"
SELECT
    si.id,
    si.document_guid,
    si.document_name,
    si.folder_path,
    LEN(LTRIM(RTRIM(si.document_guid))) AS guid_len,
    CASE WHEN EXISTS (
        SELECT 1 FROM sheet_index si2
        WHERE si2.folder_path = si.folder_path
          AND si2.document_name = si.document_name
          AND si2.document_guid <> si.document_guid
          AND LEN(LTRIM(RTRIM(si2.document_guid))) >= @minLen
    ) THEN 1 ELSE 0 END AS has_valid_sibling
FROM sheet_index si
WHERE LEN(LTRIM(RTRIM(ISNULL(si.document_guid, '')))) < @minLen
$folderClause
ORDER BY si.folder_path, si.document_name, si.id
"@

$preview = Invoke-QCDatabaseQuery -Config $config -Sql $previewSql -Parameters $params
$rows = @()
if ($preview.Data.table) {
    foreach ($r in $preview.Data.table.Rows) {
        $rows += [pscustomobject]@{
            id              = [int]$r.id
            documentGuid    = [string]$r.document_guid
            documentName    = [string]$r.document_name
            folderPath      = [string]$r.folder_path
            guidLen         = [int]$r.guid_len
            hasValidSibling = ([int]$r.has_valid_sibling) -eq 1
        }
    }
}

$summary = [ordered]@{
    minGuidLength   = $MinGuidLength
    folderFilter    = $FolderPathFilter
    rowsMatched     = $rows.Count
    withValidSibling = @($rows | Where-Object { $_.hasValidSibling }).Count
    orphanShortGuids = @($rows | Where-Object { -not $_.hasValidSibling }).Count
    deleted         = 0
    dryRun          = $DryRun.IsPresent
    samples         = @($rows | Select-Object -First 25)
}

if ($rows.Count -eq 0) {
    Write-Host 'No sheet_index rows with document_guid shorter than MinGuidLength.' -ForegroundColor Green
    if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }
    return
}

Write-Host ("Found {0} row(s) with document_guid length < {1}." -f $rows.Count, $MinGuidLength) -ForegroundColor Cyan
Write-Host ("  {0} have a sibling row with a longer guid (import duplicates)." -f $summary.withValidSibling)
if ($summary.orphanShortGuids -gt 0) {
    Write-Host ("  {0} have no sibling with a longer guid (orphans; still removed)." -f $summary.orphanShortGuids) -ForegroundColor Yellow
}

foreach ($row in @($rows | Select-Object -First 15)) {
    $tag = if ($row.hasValidSibling) { 'dup' } else { 'orphan' }
    Write-Host ("  [{0}] id={1} guid={2} len={3} | {4}" -f $tag, $row.id, $row.documentGuid, $row.guidLen, $row.documentName) -ForegroundColor Gray
}
if ($rows.Count -gt 15) {
    Write-Host ("  ... and {0} more (use -Pretty for full list)" -f ($rows.Count - 15)) -ForegroundColor Gray
}

if (-not $doDelete) {
    Write-Host 'Dry run: no rows deleted. Pass -ConfirmDeletes to remove them.' -ForegroundColor Yellow
    if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
    return
}

$ids = @($rows | ForEach-Object { $_.id })
$deleted = 0
$chunkSize = 200
for ($i = 0; $i -lt $ids.Count; $i += $chunkSize) {
    $chunk = @($ids[$i..[Math]::Min($i + $chunkSize - 1, $ids.Count - 1)])
    $delParams = @{ minLen = $MinGuidLength }
    $inList = New-Object System.Collections.Generic.List[string]
    for ($j = 0; $j -lt $chunk.Count; $j++) {
        $key = "id$j"
        $delParams[$key] = $chunk[$j]
        $inList.Add("@$key") | Out-Null
    }
    $idClause = ($inList -join ', ')
    # Delete by previewed ids only; folder filter already applied in SELECT.
    $deleteSql = @"
DELETE FROM sheet_index
WHERE id IN ($idClause)
  AND LEN(LTRIM(RTRIM(ISNULL(document_guid, '')))) < @minLen
"@
    if ($PSCmdlet.ShouldProcess(("{0} sheet_index row(s)" -f $chunk.Count), 'DELETE invalid document_guid rows')) {
        $delRes = Invoke-QCDatabaseNonQuery -Config $config -Sql $deleteSql -Parameters $delParams
        if (-not $delRes.IsSuccess) { throw $delRes.Message }
        $deleted += [int]$delRes.Data.rowsAffected
    }
}

$summary.deleted = $deleted
Write-Host ("Deleted {0} row(s) from sheet_index (v_sheet_status reflects the change)." -f $deleted) -ForegroundColor Green
if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
