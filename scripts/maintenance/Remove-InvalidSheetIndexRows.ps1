<#
.SYNOPSIS
Removes invalid rows from sheet_index (short GUIDs and/or paths outside CADD/Sheets).

.DESCRIPTION
Two common bad-data patterns:
1. ShortGuid — early imports stored document_number in document_guid (e.g. "258").
2. OutsideSheets — rows whose folder_path is not under ...\CADD\Sheets.

Deletes from dbo.sheet_index only. v_sheet_status is a view over that table.

Default is preview. Pass -ConfirmDeletes to apply.

.EXAMPLE
.\scripts\Remove-InvalidSheetIndexRows.ps1 -DryRun -Target All
.\scripts\Remove-InvalidSheetIndexRows.ps1 -ConfirmDeletes -Target OutsideSheets
.\scripts\Remove-InvalidSheetIndexRows.ps1 -ConfirmDeletes -Target ShortGuid -FolderPathFilter 'AZFWY2302-018'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [ValidateSet('ShortGuid', 'OutsideSheets', 'All')]
    [string]$Target = 'All',
    [string]$FolderPathFilter = '',
    [int]$MinGuidLength = 5,
    [string]$SheetsPathFragment = 'CADD/Sheets',
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
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

$sheetsLike = '%{0}%' -f (($SheetsPathFragment -replace '\\', '/').ToLowerInvariant())

function _RSI-BuildFolderFilterClause {
    param([hashtable]$Params)
    if ([string]::IsNullOrWhiteSpace($FolderPathFilter)) { return '' }
    $Params['folderLike'] = ('%{0}%' -f $FolderPathFilter.Trim())
    return ' AND folder_path LIKE @folderLike'
}

function _RSI-DeleteRowsById {
    param(
        [hashtable]$Config,
        [int[]]$Ids,
        [string]$ReasonLabel,
        [hashtable]$ExtraDeleteParams = @{}
    )
    if (-not $Ids -or $Ids.Count -eq 0) { return 0 }

    $deleted = 0
    $chunkSize = 200
    for ($i = 0; $i -lt $Ids.Count; $i += $chunkSize) {
        $chunk = @($Ids[$i..[Math]::Min($i + $chunkSize - 1, $Ids.Count - 1)])
        $delParams = @{} + $ExtraDeleteParams
        $inList = New-Object System.Collections.Generic.List[string]
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $key = "id$j"
            $delParams[$key] = $chunk[$j]
            $inList.Add("@$key") | Out-Null
        }
        $idClause = ($inList -join ', ')
        $deleteSql = "DELETE FROM sheet_index WHERE id IN ($idClause)"
        if ($PSCmdlet.ShouldProcess(("{0} sheet_index row(s) [{1}]" -f $chunk.Count, $ReasonLabel), 'DELETE')) {
            $delRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $deleteSql -Parameters $delParams
            if (-not $delRes.IsSuccess) { throw $delRes.Message }
            $deleted += [int]$delRes.Data.rowsAffected
        }
    }
    return $deleted
}

function _RSI-RunShortGuidCleanup {
    param([hashtable]$Config)

    $params = @{ minLen = $MinGuidLength }
    $folderClause = _RSI-BuildFolderFilterClause -Params $params
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

    $preview = Invoke-QCDatabaseQuery -Config $Config -Sql $previewSql -Parameters $params
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

    $result = [ordered]@{
        target           = 'ShortGuid'
        rowsMatched      = $rows.Count
        withValidSibling = @($rows | Where-Object { $_.hasValidSibling }).Count
        orphanShortGuids = @($rows | Where-Object { -not $_.hasValidSibling }).Count
        deleted          = 0
        samples          = @($rows | Select-Object -First 15)
    }

    if ($rows.Count -eq 0) {
        Write-Host '[ShortGuid] No rows with document_guid shorter than MinGuidLength.' -ForegroundColor Green
        return $result
    }

    Write-Host ("[ShortGuid] Found {0} row(s) with document_guid length < {1}." -f $rows.Count, $MinGuidLength) -ForegroundColor Cyan
    foreach ($row in @($rows | Select-Object -First 15)) {
        $tag = if ($row.hasValidSibling) { 'dup' } else { 'orphan' }
        Write-Host ("  [{0}] id={1} guid={2} | {3}" -f $tag, $row.id, $row.documentGuid, $row.documentName) -ForegroundColor Gray
    }
    if ($rows.Count -gt 15) {
        Write-Host ("  ... and {0} more" -f ($rows.Count - 15)) -ForegroundColor Gray
    }

    if ($doDelete) {
        $result.deleted = _RSI-DeleteRowsById -Config $Config -Ids @($rows | ForEach-Object { $_.id }) -ReasonLabel 'ShortGuid'
        Write-Host ("[ShortGuid] Deleted {0} row(s)." -f $result.deleted) -ForegroundColor Green
    }
    return $result
}

function _RSI-RunOutsideSheetsCleanup {
    param([hashtable]$Config)

    $params = @{ sheetsLike = $sheetsLike }
    $folderClause = _RSI-BuildFolderFilterClause -Params $params
    $previewSql = @"
SELECT id, document_guid, document_name, folder_path
FROM sheet_index
WHERE REPLACE(LOWER(LTRIM(RTRIM(ISNULL(folder_path, '')))), '\', '/') NOT LIKE @sheetsLike
$folderClause
ORDER BY folder_path, document_name, id
"@

    $preview = Invoke-QCDatabaseQuery -Config $Config -Sql $previewSql -Parameters $params
    $rows = @()
    if ($preview.Data.table) {
        foreach ($r in $preview.Data.table.Rows) {
            $rows += [pscustomobject]@{
                id           = [int]$r.id
                documentGuid = [string]$r.document_guid
                documentName = [string]$r.document_name
                folderPath   = [string]$r.folder_path
            }
        }
    }

    $result = [ordered]@{
        target          = 'OutsideSheets'
        sheetsFragment  = $SheetsPathFragment
        rowsMatched     = $rows.Count
        deleted         = 0
        samples         = @($rows | Select-Object -First 15)
    }

    if ($rows.Count -eq 0) {
        Write-Host ("[OutsideSheets] No rows outside {0}." -f $SheetsPathFragment) -ForegroundColor Green
        return $result
    }

    Write-Host ("[OutsideSheets] Found {0} row(s) not under {1}." -f $rows.Count, $SheetsPathFragment) -ForegroundColor Cyan
    foreach ($row in @($rows | Select-Object -First 15)) {
        Write-Host ("  id={0} | {1} | {2}" -f $row.id, $row.documentName, $row.folderPath) -ForegroundColor Gray
    }
    if ($rows.Count -gt 15) {
        Write-Host ("  ... and {0} more" -f ($rows.Count - 15)) -ForegroundColor Gray
    }

    if ($doDelete) {
        $result.deleted = _RSI-DeleteRowsById -Config $Config -Ids @($rows | ForEach-Object { $_.id }) -ReasonLabel 'OutsideSheets'
        Write-Host ("[OutsideSheets] Deleted {0} row(s)." -f $result.deleted) -ForegroundColor Green
    }
    return $result
}

$summary = [ordered]@{
    target         = $Target
    folderFilter   = $FolderPathFilter
    dryRun         = $DryRun.IsPresent
    shortGuid      = $null
    outsideSheets  = $null
    totalDeleted   = 0
}

$targets = if ($Target -eq 'All') { @('ShortGuid', 'OutsideSheets') } else { @($Target) }
foreach ($t in $targets) {
    if ($t -eq 'ShortGuid') {
        $summary.shortGuid = _RSI-RunShortGuidCleanup -Config $config
        if ($doDelete) { $summary.totalDeleted += [int]$summary.shortGuid.deleted }
    } elseif ($t -eq 'OutsideSheets') {
        $summary.outsideSheets = _RSI-RunOutsideSheetsCleanup -Config $config
        if ($doDelete) { $summary.totalDeleted += [int]$summary.outsideSheets.deleted }
    }
}

if (-not $doDelete) {
    Write-Host 'Dry run: no rows deleted. Pass -ConfirmDeletes to remove matched rows.' -ForegroundColor Yellow
} elseif ($summary.totalDeleted -gt 0) {
    Write-Host ("Total deleted: {0} (v_sheet_status reflects changes)." -f $summary.totalDeleted) -ForegroundColor Green
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 8 }
