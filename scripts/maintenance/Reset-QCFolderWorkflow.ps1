<#
.SYNOPSIS
Resets ProjectWise workflow state and/or clears QC telemetry in SQL by folder, stem PDF path, or QC PDF path.

.DESCRIPTION
For a Sheets folder, one stem PDF/DGN path, or one lane QC PDF path (*-prod/-chk/-rev.pdf):

1. ProjectWise - sets workflow state on documents in scope to the target state
   (default: qcWorkflow.states.production, usually "In Production").
2. Database - deletes scoped rows from QC telemetry tables (not sheet_index, sheet_packages, or sheet_documents).

Double-click the script (or run with no path parameters) for an interactive menu. Option 1 lists folder
paths from the database so you can pick all, specific numbers (e.g. 1,3 or 1-3), or type a path manually.
Interactive mode always skips ProjectWise and applies database changes only after you confirm.

Default is preview only. Pass -ConfirmReset to apply PW and database changes.

sheet_index, sheet_packages, and sheet_documents rows are never deleted by default, except lane QC PDF
registry cleanup: *-prod.pdf / *-chk.pdf / *-rev.pdf sheet_index rows are DELETED (ghost GUID cleanup after
manual lane PDF deletes), and sheet_documents rows with document_role = qc_pdf are DELETED. Pass
-KeepLanePdfRegistry to preserve lane index rows (legacy UPDATE-only behavior).

By default the script updates pw_state_name on remaining sheet_index rows (stem PDF, DGN, etc.) and clears
qc_stage/qc_status on sheet_index unless -KeepSheetIndexQcFields), clears qc_process_type and lane QC PDF pairing
(qc_pdf_guid/name on sheet_index; qc_pdf_*, qc_chk_pdf_*, qc_rev_pdf_* on sheet_packages), deletes sheet_package_qc_pdfs
rows for matched packages, clears qc_cycle_id/qc_cycle_number, zeros production_qc_completed_count,
production_qc_last_completed_at, peer_review_completed_count, peer_review_last_completed_at,
independent_check_completed_count, and independent_check_last_completed_at on sheet_index and sheet_packages,
clears qc_review_type/qc_assigned_to on sheet_index (unless -KeepSheetIndexQcFields) and sheet_packages,
updates sheet_documents.pw_state_name, and deletes qc_cycle_completions rows for the folder
(by document_guid or sheet_package_id). Pass -SkipSheetIndexUpdate to leave sheet_index completely unchanged.
Does not remove queue JSON jobs (use Purge-QCPendingByFilters or manual queue cleanup separately).

Does not delete ProjectWise documents, sheet_packages rows, or non-lane sheet_index / sheet_documents rows.

Recommended lane-PDF recycle workflow:
1. Manually delete *-prod.pdf / *-chk.pdf / *-rev.pdf in ProjectWise (DOCUMENT_DELETE registry cleanup runs via watcher).
2. Run this script with -ConfirmReset to purge telemetry and lane registry ghosts, reset stem/DGN states.
3. Restart QC prepend cycle from Initiate Origination on the stem PDF.

.EXAMPLE
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath 'AZFWY1704-FD02-SR202\CADD\Sheets' -ConfirmReset
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -StemPath 'Documents\...\Sheets\Seg_1\080j082001ab001.pdf' -ConfirmReset -SkipProjectWise
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -QcPdfPath 'Documents\...\Sheets\Seg_1\080j082001ab001-prod.pdf' -ConfirmReset -SkipProjectWise
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipProjectWise
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipDatabase
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipSheetIndexUpdate
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -KeepLanePdfRegistry
.\scripts\maintenance\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipPreviewCounts -QueryTimeoutSeconds 600

Double-click (or run with no path parameters) for an interactive menu: folder, stem PDF path, or QC PDF path.
Interactive mode always skips ProjectWise and applies database reset only after confirmation.
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Folder')]
param(
    [Parameter(ParameterSetName = 'Folder')]
    [string]$FolderPath = '',

    [Parameter(ParameterSetName = 'Stem')]
    [string]$StemPath = '',

    [Parameter(ParameterSetName = 'QcPdf')]
    [string]$QcPdfPath = '',

    [ValidateSet('Folder', 'Stem', 'QcPdf')]
    [string]$ScopeMode = '',

    [string]$AppSettingsPath = '',
    [string]$TargetState = '',
    [switch]$ConfirmReset,
    [switch]$DryRun,
    [switch]$SkipProjectWise,
    [switch]$SkipDatabase,
    [switch]$IncludeCommentTelemetry,
    [switch]$KeepSheetIndexQcFields,
    [switch]$SkipSheetIndexUpdate,
    [switch]$KeepLanePdfRegistry,
    [int]$BatchSize = 5000,
    [int]$QueryTimeoutSeconds = 300,
    [switch]$SkipPreviewCounts,
    [switch]$Pretty,
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Database\Core.Database.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.Discovery.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Test-QCDatabaseEnabled'
    'New-QCDatabaseSession'
) -Context 'Reset-QCFolderWorkflow bootstrap'
Import-QCModuleGlobal -RelativePath 'Core\Core.Paths.psm1'
Test-QCRequiredCommands -Names @('Normalize-QCDocumentsFolderPath', 'Split-QCPathParts') -Context 'Reset-QCFolderWorkflow paths post-restore'

function _RQCF-PauseIfInteractiveConsole {
    if ($Host.Name -ne 'ConsoleHost') { return }
    try {
        Write-Host ''
        Write-Host 'Press Enter to close this window...' -ForegroundColor Yellow
        $null = Read-Host
    } catch { }
}

function _RQCF-ParseFolderSelectionInput {
    param(
        [Parameter(Mandatory)][string]$SelectionText,
        [Parameter(Mandatory)][string[]]$AvailablePaths
    )
    $text = $SelectionText.Trim()
    if ($text -match '^(?i)(all|\*|a)$') {
        return @($AvailablePaths)
    }
    if ($text -match '[\\/]' -or $text -imatch '^documents') {
        $normRes = Normalize-QCDocumentsFolderPath -Path $text
        if (-not $normRes.IsSuccess) { throw $normRes.Message }
        return @([string]$normRes.Data.path)
    }

    $selected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($part in ($text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if ($part -match '^(\d+)\s*-\s*(\d+)$') {
            $start = [int]$Matches[1]
            $end = [int]$Matches[2]
            if ($start -gt $end) { throw "Invalid range: $part" }
            for ($n = $start; $n -le $end; $n++) {
                if ($n -lt 1 -or $n -gt $AvailablePaths.Count) {
                    throw "Selection out of range: $n (1-$($AvailablePaths.Count))"
                }
                [void]$selected.Add($AvailablePaths[$n - 1])
            }
            continue
        }
        if ($part -match '^\d+$') {
            $n = [int]$part
            if ($n -lt 1 -or $n -gt $AvailablePaths.Count) {
                throw "Selection out of range: $n (1-$($AvailablePaths.Count))"
            }
            [void]$selected.Add($AvailablePaths[$n - 1])
            continue
        }
        throw "Unrecognized selection '$part'. Use all, numbers (1,3,5 or 17-23), or a folder path."
    }
    if ($selected.Count -eq 0) { throw 'No folders selected.' }
    return @($selected | Sort-Object)
}

function _RQCF-GetAvailableFolderPathsFromDatabase {
    param([Parameter(Mandatory)][hashtable]$Config)
    Initialize-QCDatabaseSchema -Config $Config | Out-Null
    $sql = @"
SELECT f.folder_path,
       ISNULL(p.package_count, 0) AS package_count,
       ISNULL(i.index_count, 0) AS index_count
FROM (
    SELECT folder_path FROM sheet_packages WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> ''
    UNION
    SELECT folder_path FROM sheet_index WHERE folder_path IS NOT NULL AND LTRIM(RTRIM(folder_path)) <> ''
) f
LEFT JOIN (
    SELECT folder_path, COUNT_BIG(*) AS package_count FROM sheet_packages GROUP BY folder_path
) p ON p.folder_path = f.folder_path
LEFT JOIN (
    SELECT folder_path, COUNT_BIG(*) AS index_count FROM sheet_index GROUP BY folder_path
) i ON i.folder_path = f.folder_path
ORDER BY f.folder_path
"@
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql
    if (-not $res.IsSuccess) { throw $res.Message }
    $table = $res.Data.table
    if (-not $table -or $table.Rows.Count -eq 0) { return @() }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($dataRow in @($table.Rows)) {
        $rows.Add([pscustomobject]@{
            folder_path    = [string]$dataRow.folder_path
            package_count  = [long]$dataRow.package_count
            index_count    = [long]$dataRow.index_count
        }) | Out-Null
    }
    return @($rows)
}

function _RQCF-PromptInteractiveFolderSelection {
    param([Parameter(Mandatory)][hashtable]$Config)
    Write-Host ''
    Write-Host 'Loading folder paths from database...' -ForegroundColor Gray
    $available = @(_RQCF-GetAvailableFolderPathsFromDatabase -Config $Config)
    if ($available.Count -eq 0) {
        Write-Host 'No folder paths found in database.' -ForegroundColor Yellow
        return $null
    }

    $paths = [System.Collections.Generic.List[string]]::new()
    Write-Host ''
    Write-Host 'Available folder paths:' -ForegroundColor Cyan
    $idx = 1
    foreach ($row in $available) {
        $paths.Add([string]$row.folder_path) | Out-Null
        Write-Host ("  {0,4}) {1}  (packages: {2}, index: {3})" -f $idx, $row.folder_path, $row.package_count, $row.index_count)
        $idx++
    }
    Write-Host ''
    Write-Host "Enter selection: all ($($paths.Count) folders), numbers (e.g. 1,3,5 or 1-3), or type a folder path." -ForegroundColor DarkGray
    do {
        $raw = (Read-Host 'Folder selection').Trim()
    } while ([string]::IsNullOrWhiteSpace($raw))

    $selected = @(_RQCF-ParseFolderSelectionInput -SelectionText $raw -AvailablePaths @($paths))
    $manualPath = ($raw -match '[\\/]' -or $raw -imatch '^documents')
    return @{
        paths                    = @($selected)
        folderPickFromDatabase   = -not $manualPath
    }
}

function _RQCF-PromptInteractiveScope {
    param([hashtable]$Config = $null)

    Write-Host ''
    Write-Host '=== QC Database Reset ===' -ForegroundColor Cyan
    Write-Host 'ProjectWise is skipped in interactive mode (database edits only).' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  1) Reset by folder path'
    Write-Host '  2) Reset by stem PDF path (sheet stem + DGN + lane PDFs for one sheet)'
    Write-Host '  3) Reset by QC PDF path (one lane PDF: *-prod/-chk/-rev.pdf)'
    Write-Host ''
    do {
        $choice = (Read-Host 'Select reset scope (1/2/3)').Trim()
    } while ($choice -notin @('1', '2', '3'))

    if ($choice -eq '1' -and $Config -and (Test-QCDatabaseEnabled -Config $Config)) {
        try {
            $folderPick = _RQCF-PromptInteractiveFolderSelection -Config $Config
            if ($folderPick) {
                return @{
                    scopeMode              = 'Folder'
                    inputPath              = ($folderPick.paths -join '; ')
                    folderPaths            = @($folderPick.paths)
                    folderPickFromDatabase = [bool]$folderPick.folderPickFromDatabase
                }
            }
        } catch {
            Write-Host ("Could not list folders from database: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            Write-Host 'Enter folder path manually.' -ForegroundColor DarkGray
        }
    }

    $prompt = switch ($choice) {
        '1' { @('Enter folder path (Documents\...\Sheets\...)', 'Folder') }
        '2' { @('Enter stem PDF path (Documents\...\sheetstem.pdf or .dgn)', 'Stem') }
        '3' { @('Enter QC PDF path (Documents\...\sheetstem-prod.pdf)', 'QcPdf') }
    }
    $label = $prompt[1]
    $path = ''
    do {
        $path = (Read-Host $prompt[0]).Trim()
    } while ([string]::IsNullOrWhiteSpace($path))
    return @{
        scopeMode              = $label
        inputPath              = $path
        folderPaths            = @()
        folderPickFromDatabase = $false
    }
}

function _RQCF-ParsePwDocumentPath {
    param([Parameter(Mandatory)][string]$Path)
    $splitRes = Split-QCPathParts -Path $Path
    if (-not $splitRes.IsSuccess) { throw $splitRes.Message }
    $full = [string]$splitRes.Data.normalizedPath
    $leaf = [string]$splitRes.Data.leaf
    if ([string]::IsNullOrWhiteSpace($leaf)) { throw 'Path must include a document file name.' }
    $parent = if ($full.Length -gt $leaf.Length) { $full.Substring(0, $full.Length - $leaf.Length).TrimEnd('\') } else { '' }
    if ([string]::IsNullOrWhiteSpace($parent)) { throw 'Path must include a folder and document file name.' }
    $folderRes = Normalize-QCDocumentsFolderPath -Path $parent
    if (-not $folderRes.IsSuccess) { throw $folderRes.Message }
    return @{
        folderPath   = [string]$folderRes.Data.path
        documentName = $leaf
        fullPath     = $full
    }
}

function _RQCF-ResolveSheetStemFromDocumentName {
    param([Parameter(Mandatory)][string]$DocumentName)
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $stem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
        if (-not [string]::IsNullOrWhiteSpace($stem)) { return $stem }
    }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ($stem -match '(?i)-(prod|chk|rev)$') { $stem = $stem -replace '(?i)-(prod|chk|rev)$', '' }
    return $stem
}

function _RQCF-GetLaneTypeFromDocumentName {
    param([Parameter(Mandatory)][string]$DocumentName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ($base -match '(?i)-(prod)$') { return 'production' }
    if ($base -match '(?i)-(chk)$') { return 'check' }
    if ($base -match '(?i)-(rev)$') { return 'review' }
    return ''
}

$script:_rqcfInteractiveLaunch = $false
$script:_rqcfInteractiveFolderPick = $false
$config = $null
if ($Interactive.IsPresent) {
    $script:_rqcfInteractiveLaunch = $true
} elseif ($Host.Name -eq 'ConsoleHost' -and
    [string]::IsNullOrWhiteSpace($FolderPath) -and
    [string]::IsNullOrWhiteSpace($StemPath) -and
    [string]::IsNullOrWhiteSpace($QcPdfPath)) {
    $script:_rqcfInteractiveLaunch = $true
}

if ($script:_rqcfInteractiveLaunch) {
    $cfgResEarly = Read-QCAppSettings -Path $AppSettingsPath
    if (-not $cfgResEarly.IsSuccess) { throw $cfgResEarly.Message }
    $config = [hashtable]$cfgResEarly.Data.config
    $prompted = _RQCF-PromptInteractiveScope -Config $config
    $ScopeMode = [string]$prompted.scopeMode
    switch ($ScopeMode) {
        'Folder' {
            $FolderPath = [string]$prompted.inputPath
            if (@($prompted.folderPaths).Count -gt 0) {
                $script:_rqcfInteractiveFolderPick = [bool]$prompted.folderPickFromDatabase
            }
        }
        'Stem' { $StemPath = [string]$prompted.inputPath }
        'QcPdf' { $QcPdfPath = [string]$prompted.inputPath }
    }
    $SkipProjectWise = $true
}

if ([string]::IsNullOrWhiteSpace($ScopeMode)) {
    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) { $ScopeMode = 'Folder' }
    elseif (-not [string]::IsNullOrWhiteSpace($StemPath)) { $ScopeMode = 'Stem' }
    elseif (-not [string]::IsNullOrWhiteSpace($QcPdfPath)) { $ScopeMode = 'QcPdf' }
    else { throw 'Pass -FolderPath, -StemPath, -QcPdfPath, or double-click for interactive mode.' }
}

$scopeInputPath = switch ($ScopeMode) {
    'Folder' { $FolderPath }
    'Stem' { $StemPath }
    'QcPdf' { $QcPdfPath }
    default { throw "Unsupported ScopeMode: $ScopeMode" }
}
if ([string]::IsNullOrWhiteSpace($scopeInputPath)) {
    throw "ScopeMode '$ScopeMode' requires a path parameter."
}

$scopeUsesFolder = ($ScopeMode -eq 'Folder')
$scopeStem = ''
$scopeDocName = ''
$scopeDocPathLike = ''
$normFolder = ''
$normFolders = @()

if ($scopeUsesFolder) {
    if ($script:_rqcfInteractiveLaunch -and @($prompted.folderPaths).Count -gt 0) {
        foreach ($folderPath in @($prompted.folderPaths)) {
            $normRes = Normalize-QCDocumentsFolderPath -Path $folderPath
            if (-not $normRes.IsSuccess) { throw $normRes.Message }
            $normFolders += [string]$normRes.Data.path
        }
    } else {
        $normRes = Normalize-QCDocumentsFolderPath -Path $scopeInputPath
        if (-not $normRes.IsSuccess) { throw $normRes.Message }
        $normFolders = @([string]$normRes.Data.path)
    }
    if ($normFolders.Count -eq 1) {
        $normFolder = $normFolders[0]
    } else {
        $normFolder = "$($normFolders.Count) folders"
    }
} else {
    $parsed = _RQCF-ParsePwDocumentPath -Path $scopeInputPath
    $normFolder = [string]$parsed.folderPath
    $scopeDocName = [string]$parsed.documentName
    if ($ScopeMode -eq 'Stem') {
        $scopeStem = _RQCF-ResolveSheetStemFromDocumentName -DocumentName $scopeDocName
        if ([string]::IsNullOrWhiteSpace($scopeStem)) { throw "Could not derive sheet stem from path: $scopeInputPath" }
    } elseif ($ScopeMode -eq 'QcPdf') {
        if ($scopeDocName -notmatch '(?i)-(prod|chk|rev)\.pdf$') {
            throw 'QC PDF path must end with -prod.pdf, -chk.pdf, or -rev.pdf.'
        }
        $scopeStem = _RQCF-ResolveSheetStemFromDocumentName -DocumentName $scopeDocName
    }
}

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if ($SkipProjectWise.IsPresent) { break }
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

if ($DryRun.IsPresent -and $ConfirmReset.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmReset (apply changes), not both.'
}
$doApply = $ConfirmReset.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmReset.IsPresent) {
    if ($script:_rqcfInteractiveLaunch) {
        Write-Host 'Preview only (database). You will be asked to confirm before applying.' -ForegroundColor Yellow
    } else {
        Write-Host 'Preview only: pass -ConfirmReset to apply PW + database reset, or -DryRun explicitly.' -ForegroundColor Yellow
    }
    $DryRun = $true
}

$cfgRes = if ($null -ne $config) {
    New-QCSuccessResult -Code 'CONFIG_CACHED' -Message 'Using config loaded for interactive mode.' -Data @{ config = $config }
} else {
    Read-QCAppSettings -Path $AppSettingsPath
}
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not $SkipDatabase.IsPresent) {
    if (-not (Test-QCDatabaseEnabled -Config $config)) {
        throw 'database.enabled must be true (or pass -SkipDatabase).'
    }
    Initialize-QCDatabaseSchema -Config $config | Out-Null
}

if ([string]::IsNullOrWhiteSpace($TargetState)) {
    $TargetState = 'In Production'
    try {
        if ($config.ContainsKey('qcWorkflow') -and $config.qcWorkflow -and $config.qcWorkflow.states -and $config.qcWorkflow.states.production) {
            $TargetState = [string]$config.qcWorkflow.states.production
        }
    } catch { }
}
if ([string]::IsNullOrWhiteSpace($TargetState)) {
    throw 'TargetState is empty; set qcWorkflow.states.production in appsettings or pass -TargetState.'
}

function _RQCF-NormalizePathPattern {
    param([string]$Pattern)
    $norm = ($Pattern -replace '\\', '/').ToLowerInvariant()
    if ($norm -notmatch '[%_\[]') { return "%$norm%" }
    return $norm
}

function _RQCF-PathLikeClause {
    param(
        [string]$ColumnSql,
        [hashtable]$Params,
        [string]$ParamName = 'folderLike'
    )
    return "REPLACE(LOWER(LTRIM(RTRIM(ISNULL($ColumnSql, '')))), '\', '/') LIKE @$ParamName"
}

function _RQCF-BuildFolderScopeMatchClause {
    param(
        [Parameter(Mandatory)][string]$ColumnSql,
        [Parameter(Mandatory)][string[]]$FolderPaths,
        [Parameter(Mandatory)][hashtable]$Params,
        [switch]$UseLikePattern
    )
    if ($UseLikePattern -and $FolderPaths.Count -eq 1) {
        $Params.folderLike = _RQCF-NormalizePathPattern -Pattern $FolderPaths[0]
        return _RQCF-PathLikeClause -ColumnSql $ColumnSql -Params $Params -ParamName 'folderLike'
    }
    if ($FolderPaths.Count -eq 0) { throw 'No folder paths in scope.' }

    $parts = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $FolderPaths.Count; $i++) {
        $paramName = "scopeFolder$i"
        $Params[$paramName] = $FolderPaths[$i]
        $normCol = "REPLACE(LOWER(LTRIM(RTRIM(ISNULL($ColumnSql, '')))), '\', '/')"
        $normParam = "REPLACE(LOWER(LTRIM(RTRIM(ISNULL(@$paramName, '')))), '\', '/')"
        [void]$parts.Add("$normCol = $normParam")
    }
    if ($parts.Count -eq 1) { return $parts[0] }
    return '(' + ($parts -join ' OR ') + ')'
}

function _RQCF-InPackageScope {
    param(
        [Parameter(Mandatory)][string]$PackageIdColumn,
        [Parameter(Mandatory)][string]$SheetPackageFolderClause
    )
    return "EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = $PackageIdColumn AND ($SheetPackageFolderClause))"
}

function _RQCF-InDocumentScope {
    param(
        [Parameter(Mandatory)][string]$DocumentGuidColumn,
        [Parameter(Mandatory)][string]$SheetIndexFolderClause,
        [Parameter(Mandatory)][string]$SheetPackageFolderClause
    )
    return @"
(
    EXISTS (
        SELECT 1
        FROM sheet_documents sd
        INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
        WHERE ($SheetPackageFolderClause)
          AND LTRIM(RTRIM(CAST(sd.document_guid AS NVARCHAR(50)))) = LTRIM(RTRIM($DocumentGuidColumn))
    )
    OR EXISTS (
        SELECT 1 FROM sheet_index si
        WHERE ($SheetIndexFolderClause)
          AND LTRIM(RTRIM(si.document_guid)) = LTRIM(RTRIM($DocumentGuidColumn))
    )
    OR EXISTS (
        SELECT 1 FROM sheet_index si
        WHERE ($SheetIndexFolderClause)
          AND LTRIM(RTRIM(si.qc_pdf_guid)) = LTRIM(RTRIM($DocumentGuidColumn))
    )
)
"@
}

function _RQCF-InStemSheetIndexScope {
    return @"
(
    EXISTS (
        SELECT 1 FROM sheet_packages sp0
        WHERE sp0.folder_path = @scopeFolder AND sp0.sheet_stem = @scopeStem
          AND si.sheet_package_id = sp0.sheet_package_id
    )
    OR (
        si.folder_path = @scopeFolder
        AND (
            LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemPdfName)))
            OR LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemDgnName)))
            OR LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemProdName)))
            OR LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemChkName)))
            OR LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemRevName)))
        )
    )
)
"@
}

function _RQCF-TelemetryScopeClause {
    param(
        [string]$FolderClause = '',
        [string]$PackageIdColumn = '',
        [string]$DocumentGuidColumn = '',
        [Parameter(Mandatory)][string]$SheetIndexFolderClause,
        [Parameter(Mandatory)][string]$SheetPackageFolderClause
    )
    $parts = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($FolderClause)) { [void]$parts.Add("($FolderClause)") }
    if (-not [string]::IsNullOrWhiteSpace($PackageIdColumn)) {
        [void]$parts.Add((_RQCF-InPackageScope -PackageIdColumn $PackageIdColumn -SheetPackageFolderClause $SheetPackageFolderClause))
    }
    if (-not [string]::IsNullOrWhiteSpace($DocumentGuidColumn)) {
        [void]$parts.Add((_RQCF-InDocumentScope -DocumentGuidColumn $DocumentGuidColumn -SheetIndexFolderClause $SheetIndexFolderClause -SheetPackageFolderClause $SheetPackageFolderClause))
    }
    return ($parts -join ' OR ')
}

function _RQCF-GetScalarConn {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql,
        [hashtable]$Params,
        [int]$CommandTimeout
    )
    $res = Invoke-QCDatabaseScalarWithConnection -Connection $Connection -Sql $Sql -Parameters $Params -CommandTimeout $CommandTimeout
    if (-not $res.IsSuccess) { throw $res.Message }
    return [long]$res.Data.value
}

function _RQCF-RunDeleteLoopConn {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql,
        [hashtable]$Params,
        [string]$Label,
        [int]$CommandTimeout
    )
    $total = 0L
    do {
        $batch = 0
        if ($PSCmdlet.ShouldProcess($Label, 'DELETE batch')) {
            $res = Invoke-QCDatabaseNonQueryWithConnection -Connection $Connection -Sql $Sql -Parameters $Params -CommandTimeout $CommandTimeout
            if (-not $res.IsSuccess) { throw $res.Message }
            $batch = [int]$res.Data.rowsAffected
            $total += $batch
            if ($batch -gt 0) {
                Write-Host ("  [{0}] deleted {1} row(s), running total {2}" -f $Label, $batch, $total) -ForegroundColor Gray
            }
        } else {
            break
        }
    } while ($batch -ge $BatchSize)
    return $total
}

function _RQCF-RunNonQueryConn {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql,
        [hashtable]$Params,
        [int]$CommandTimeout
    )
    $res = Invoke-QCDatabaseNonQueryWithConnection -Connection $Connection -Sql $Sql -Parameters $Params -CommandTimeout $CommandTimeout
    if (-not $res.IsSuccess) { throw $res.Message }
    return [int]$res.Data.rowsAffected
}

function _RQCF-GetPwDocumentState {
    param([object]$Document)
    if (-not $Document) { return '' }
    foreach ($p in @('WorkflowState', 'StateName', 'WorkflowStateName', 'DocumentWorkflowState')) {
        try {
            if ($Document.PSObject.Properties[$p] -and $Document.$p) {
                return ([string]$Document.$p).Trim()
            }
        } catch { }
    }
    return ''
}

function _RQCF-GetLanePdfNameClause {
    param([string]$DocumentNameColumn = 'document_name')
    return @"
(
    LOWER($DocumentNameColumn) LIKE '%-prod.pdf'
    OR LOWER($DocumentNameColumn) LIKE '%-chk.pdf'
    OR LOWER($DocumentNameColumn) LIKE '%-rev.pdf'
)
"@
}

function _RQCF-SetPwDocumentState {
    param(
        [object]$Document,
        [string]$StateName
    )
    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $cmd -or -not $Document) { throw 'Set-PWDocumentState or document unavailable.' }
    $invokeArgs = @{}
    $docParam = if ($cmd.Parameters.ContainsKey('InputDocuments')) { 'InputDocuments' }
        elseif ($cmd.Parameters.ContainsKey('InputDocument')) { 'InputDocument' }
        elseif ($cmd.Parameters.ContainsKey('Document')) { 'Document' }
        else { $null }
    $stateParam = if ($cmd.Parameters.ContainsKey('StateName')) { 'StateName' }
        elseif ($cmd.Parameters.ContainsKey('State')) { 'State' }
        else { $null }
    if ($docParam) { $invokeArgs[$docParam] = @($Document) }
    if ($stateParam) { $invokeArgs[$stateParam] = $StateName }
    if ($cmd.Parameters.ContainsKey('ReturnBoolean')) { $invokeArgs['ReturnBoolean'] = $true }
    if ($docParam -and $stateParam) { & $cmd @invokeArgs -ErrorAction Stop | Out-Null }
    elseif ($stateParam) { & $cmd $Document @invokeArgs -ErrorAction Stop | Out-Null }
    else { & $cmd $Document $StateName -ErrorAction Stop | Out-Null }
}

$params = @{
    batchSize   = $BatchSize
    targetState = $TargetState
}
$scopeIsQcPdfOnly = ($ScopeMode -eq 'QcPdf')
$scopeFolderExactMatch = $script:_rqcfInteractiveFolderPick -or ($scopeUsesFolder -and $normFolders.Count -gt 1)
$scopeFolderUseLike = $scopeUsesFolder -and -not $scopeFolderExactMatch

if ($scopeUsesFolder) {
    $folderScopeArgs = @{
        FolderPaths     = $normFolders
        Params          = $params
        UseLikePattern  = $scopeFolderUseLike
    }
    $siFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'si.folder_path' @folderScopeArgs
    $siLaneNameClause = _RQCF-GetLanePdfNameClause -DocumentNameColumn 'si.document_name'
    $spFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'sp.folder_path' @folderScopeArgs
    $nlFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'n.folder_path' @folderScopeArgs
    $histFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'h.folder_path' @folderScopeArgs
    $trFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 't.folder_path' @folderScopeArgs
    $actFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'd.folder_path' @folderScopeArgs
    $auditFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'a.resolved_folder' @folderScopeArgs
    $jobFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'j.source_folder' @folderScopeArgs
} else {
    $parsedFullPath = if ($parsed) { [string]$parsed.fullPath } else { _RQCF-ParsePwDocumentPath -Path $scopeInputPath | ForEach-Object { $_.fullPath } }
    $params.scopeFolder = $normFolder
    $params.scopeStem = $scopeStem
    $params.scopeDocName = $scopeDocName
    $params.stemPdfName = "$scopeStem.pdf"
    $params.stemDgnName = "$scopeStem.dgn"
    $params.stemProdName = "$scopeStem-prod.pdf"
    $params.stemChkName = "$scopeStem-chk.pdf"
    $params.stemRevName = "$scopeStem-rev.pdf"
    $params.docPathLike = _RQCF-NormalizePathPattern -Pattern $parsedFullPath
    $scopeDocPathLike = [string]$params.docPathLike

    $spFolderClause = 'sp.folder_path = @scopeFolder AND sp.sheet_stem = @scopeStem'
    if ($scopeIsQcPdfOnly) {
        $siFolderClause = 'si.folder_path = @scopeFolder AND LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@scopeDocName)))'
        $siLaneNameClause = 'LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@scopeDocName)))'
    } else {
        $siFolderClause = _RQCF-InStemSheetIndexScope
        $siLaneNameClause = @"
(
    LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemProdName)))
    OR LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemChkName)))
    OR LOWER(LTRIM(RTRIM(si.document_name))) = LOWER(LTRIM(RTRIM(@stemRevName)))
)
"@
    }
    $nlFolderClause = ''
    $histFolderClause = ''
    $trFolderClause = ''
    $actFolderClause = ''
    $auditFolderClause = ''
    $jobFolderClause = ''
}

$scopeArgs = @{
    SheetIndexFolderClause = $siFolderClause
    SheetPackageFolderClause = $spFolderClause
}
$whereNotification = _RQCF-TelemetryScopeClause -FolderClause $nlFolderClause -PackageIdColumn 'n.sheet_package_id' -DocumentGuidColumn 'n.document_guid' @scopeArgs
$whereHistory = _RQCF-TelemetryScopeClause -FolderClause $histFolderClause -PackageIdColumn 'h.sheet_package_id' -DocumentGuidColumn 'h.document_guid' @scopeArgs
$whereTransition = _RQCF-TelemetryScopeClause -FolderClause $trFolderClause -PackageIdColumn 't.sheet_package_id' -DocumentGuidColumn 't.document_guid' @scopeArgs
$whereActivity = _RQCF-TelemetryScopeClause -FolderClause $actFolderClause -DocumentGuidColumn 'd.document_guid' @scopeArgs
$whereWorkflow = _RQCF-TelemetryScopeClause -PackageIdColumn 'w.sheet_package_id' -DocumentGuidColumn 'w.document_id' @scopeArgs
$whereCompletion = _RQCF-TelemetryScopeClause -PackageIdColumn 'c.sheet_package_id' -DocumentGuidColumn 'CAST(c.document_guid AS NVARCHAR(50))' @scopeArgs
if ($scopeUsesFolder) {
    $whereAudit = "($auditFolderClause) OR " + (_RQCF-InDocumentScope -DocumentGuidColumn 'a.pw_objguid' @scopeArgs)
} else {
    $whereAudit = _RQCF-InDocumentScope -DocumentGuidColumn 'a.pw_objguid' @scopeArgs
}
$whereJobs = _RQCF-TelemetryScopeClause -FolderClause $jobFolderClause -PackageIdColumn 'j.sheet_package_id' @scopeArgs
$whereCommentDoc = _RQCF-InDocumentScope -DocumentGuidColumn 'r.document_id' @scopeArgs
if ($scopeUsesFolder) {
    $pwPathFolderClause = _RQCF-BuildFolderScopeMatchClause -ColumnSql 'r.pw_path' -FolderPaths $normFolders -Params $params -UseLikePattern:$scopeFolderUseLike
    $commentRunSub = @"
SELECT r.run_id FROM qc_comment_runs r
WHERE ($whereCommentDoc)
   OR ($pwPathFolderClause)
"@
} else {
    $commentRunSub = @"
SELECT r.run_id FROM qc_comment_runs r
WHERE ($whereCommentDoc)
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @docPathLike
"@
}

$scopeLabel = switch ($ScopeMode) {
    'Folder' {
        if ($normFolders.Count -gt 1) {
            $preview = ($normFolders | Select-Object -First 2) -join '; '
            $suffix = if ($normFolders.Count -gt 2) { '; ...' } else { '' }
            "folders ($($normFolders.Count)): $preview$suffix"
        } else {
            "folder: $normFolder"
        }
    }
    'Stem' { "stem: $normFolder\$scopeStem" }
    'QcPdf' { "qc pdf: $normFolder\$scopeDocName" }
}
Write-Host ("[Reset scope] {0} ({1})" -f $ScopeMode, $scopeLabel) -ForegroundColor Cyan
if ($scopeUsesFolder -and $normFolders.Count -gt 1) {
    foreach ($fp in $normFolders) {
        Write-Host ("  - {0}" -f $fp) -ForegroundColor DarkGray
    }
}

$summary = [ordered]@{
    dryRun            = $DryRun.IsPresent
    scopeMode         = $ScopeMode
    folderPath        = $normFolder
    folderPaths       = @($normFolders)
    inputPath         = $scopeInputPath
    sheetStem         = $scopeStem
    documentName      = $scopeDocName
    targetState       = $TargetState
    projectWise       = $null
    database          = $null
    totalRowsDeleted  = 0
    sheetIndexUpdated = 0
    sheetPackagesUpdated = 0
    sheetDocumentsUpdated = 0
    sheetPackageQcPdfsDeleted = 0
    sheetIndexLaneRowsDeleted = 0
    sheetDocumentsQcPdfDeleted = 0
}

# --- ProjectWise ---
if (-not $SkipProjectWise.IsPresent) {
    $ds = [string]$config.projectWise.datasourceName
    $credPath = [string]$config.projectWise.credentialPath
    if ([string]::IsNullOrWhiteSpace($ds) -or [string]::IsNullOrWhiteSpace($credPath)) {
        throw 'projectWise.datasourceName and projectWise.credentialPath are required in appsettings.'
    }

    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $normFolder
    if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $normFolder }

    $script:_rqcfApiPath = $apiPath
    $script:_rqcfTargetState = $TargetState
    $script:_rqcfDoApply = $doApply

    Write-Host '[ProjectWise] Scanning folder documents...' -ForegroundColor Cyan
    try {
        $pwResult = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -ScriptBlock {
            $docs = @(Get-PWDocumentsInFolderRaw -FolderPath $script:_rqcfApiPath)
            $needsChange = 0
            $alreadyTarget = 0
            $updated = 0
            $failed = [System.Collections.Generic.List[object]]::new()
            $samples = [System.Collections.Generic.List[object]]::new()

            foreach ($doc in $docs) {
                $name = ''
                try { $name = [string]$doc.Name } catch { }
                if (-not $name) { try { $name = [string]$doc.DocumentName } catch { } }
                $cur = _RQCF-GetPwDocumentState -Document $doc
                $needs = -not ($cur -ieq $script:_rqcfTargetState)
                if ($needs) { $needsChange++ } else { $alreadyTarget++ }
                if ($samples.Count -lt 8) {
                    $samples.Add([pscustomobject]@{ name = $name; currentState = $cur; needsChange = $needs }) | Out-Null
                }
                if (-not $script:_rqcfDoApply -or -not $needs) { continue }
                try {
                    _RQCF-SetPwDocumentState -Document $doc -StateName $script:_rqcfTargetState
                    $updated++
                } catch {
                    $failed.Add([pscustomobject]@{ name = $name; error = $_.Exception.Message }) | Out-Null
                }
            }

            return [pscustomobject]@{
                documentCount = $docs.Count
                needsChange   = $needsChange
                alreadyTarget = $alreadyTarget
                updated       = $updated
                failed        = @($failed)
                samples       = @($samples)
            }
        }
    } finally {
        Remove-Variable -Name _rqcfApiPath, _rqcfTargetState, _rqcfDoApply -Scope Script -ErrorAction SilentlyContinue
    }

    $summary.projectWise = @{
        documentCount = [int]$pwResult.documentCount
        needsChange   = [int]$pwResult.needsChange
        alreadyTarget = [int]$pwResult.alreadyTarget
        updated       = [int]$pwResult.updated
        failed        = @($pwResult.failed)
        samples       = @($pwResult.samples)
    }
    Write-Host ("  Documents: {0}; to update: {1}; already '{2}': {3}" -f `
        $pwResult.documentCount, $pwResult.needsChange, $TargetState, $pwResult.alreadyTarget) -ForegroundColor Gray
    if ($doApply -and $pwResult.updated -gt 0) {
        Write-Host ("  Updated {0} document(s) to '{1}'." -f $pwResult.updated, $TargetState) -ForegroundColor Green
    }
    if ($pwResult.failed.Count -gt 0) {
        Write-Host ("  {0} document(s) failed state update." -f $pwResult.failed.Count) -ForegroundColor Yellow
        $pwResult.failed | Select-Object -First 10 | Format-Table -AutoSize
    } elseif (-not $doApply -and $pwResult.needsChange -gt 0) {
        Write-Host '  Sample documents needing update:' -ForegroundColor DarkGray
        $pwResult.samples | Where-Object { $_.needsChange } | Select-Object -First 5 | Format-Table -AutoSize
    }
} else {
    Write-Host '[ProjectWise] Skipped (-SkipProjectWise).' -ForegroundColor DarkGray
    $summary.projectWise = @{ skipped = $true }
}

# --- Database ---
if (-not $SkipDatabase.IsPresent) {
    $sessionRes = New-QCDatabaseSession -Config $config
    if (-not $sessionRes.IsSuccess) { throw $sessionRes.Message }
    $dbConn = $sessionRes.Data.session.connection
    try {
        $dbCounts = [ordered]@{}
        if (-not $SkipPreviewCounts.IsPresent) {
            Write-Host '[Database] Counting folder-scoped telemetry rows...' -ForegroundColor Cyan
            $dbCounts.notification_log = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM notification_log n WHERE ($whereNotification)
"@
            try {
                $dbCounts.qc_notification_messages = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_notification_messages m
WHERE m.notification_log_id IN (SELECT n.id FROM notification_log n WHERE ($whereNotification))
"@
            } catch { }
            $dbCounts.document_state_history = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM document_state_history h WHERE ($whereHistory)
"@
            $dbCounts.transition_events = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM transition_events t WHERE ($whereTransition)
"@
            $dbCounts.document_activity = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM document_activity d WHERE ($whereActivity)
"@
            $dbCounts.qc_workflow_events = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_workflow_events w WHERE ($whereWorkflow)
"@
            $dbCounts.qc_cycle_completions = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_cycle_completions c WHERE ($whereCompletion)
"@
            $dbCounts.audit_events = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM audit_events a WHERE ($whereAudit)
"@
            $dbCounts.processing_jobs = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM processing_jobs j WHERE ($whereJobs)
"@
            try {
                $dbCounts.sheet_package_qc_pdfs = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_package_qc_pdfs q
WHERE EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = q.sheet_package_id AND ($spFolderClause))
"@
            } catch { }
            if ($IncludeCommentTelemetry.IsPresent) {
                $whereCommentHistory = _RQCF-TelemetryScopeClause -DocumentGuidColumn 'h.document_id' @scopeArgs
                $dbCounts.qc_comment_status_history = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_comment_status_history h
WHERE ($whereCommentHistory) OR h.detected_run_id IN ($commentRunSub)
"@
                $dbCounts.qc_comments = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_comments c WHERE c.run_id IN ($commentRunSub)
"@
                $dbCounts.qc_comment_runs = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_comment_runs r
WHERE ($whereCommentDoc)
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @folderLike
"@
            }
        } else {
            Write-Host '[Database] Preview counts skipped (-SkipPreviewCounts).' -ForegroundColor DarkGray
        }

        $sheetPackagesMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_packages sp WHERE ($spFolderClause)
"@
        $sheetDocumentsMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE ($spFolderClause)
"@

        $sheetPackageQcPdfsMatched = 0L
        try {
            $sheetPackageQcPdfsMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_package_qc_pdfs q
WHERE EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = q.sheet_package_id AND ($spFolderClause))
"@
        } catch {
            $sheetPackageQcPdfsMatched = 0L
        }

        $sheetIndexMatched = 0L
        $sheetIndexLaneMatched = 0L
        $sheetIndexNonLaneMatched = 0L
        $sheetDocumentsQcPdfMatched = 0L
        $sheetIndexCompletionData = 0L
        $sheetPackageCompletionData = 0L
        if (-not $SkipSheetIndexUpdate.IsPresent) {
            $sheetIndexMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si WHERE ($siFolderClause)
"@
            if (-not $KeepLanePdfRegistry.IsPresent) {
                $sheetIndexLaneMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si WHERE ($siFolderClause) AND ($siLaneNameClause)
"@
                $sheetIndexNonLaneMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si WHERE ($siFolderClause) AND NOT ($siLaneNameClause)
"@
                $sheetDocumentsQcPdfMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE ($spFolderClause) AND sd.document_role = 'qc_pdf'
"@
            }
            try {
                $sheetIndexCompletionData = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si
WHERE ($siFolderClause)
  AND (
    si.qc_cycle_id IS NOT NULL
    OR ISNULL(si.production_qc_completed_count, 0) > 0
    OR si.production_qc_last_completed_at IS NOT NULL
    OR ISNULL(si.peer_review_completed_count, 0) > 0
    OR si.peer_review_last_completed_at IS NOT NULL
    OR ISNULL(si.independent_check_completed_count, 0) > 0
    OR si.independent_check_last_completed_at IS NOT NULL
  )
"@
                $sheetPackageCompletionData = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_packages sp
WHERE ($spFolderClause)
  AND (
    sp.qc_cycle_id IS NOT NULL
    OR ISNULL(sp.production_qc_completed_count, 0) > 0
    OR sp.production_qc_last_completed_at IS NOT NULL
    OR ISNULL(sp.peer_review_completed_count, 0) > 0
    OR sp.peer_review_last_completed_at IS NOT NULL
    OR ISNULL(sp.independent_check_completed_count, 0) > 0
    OR sp.independent_check_last_completed_at IS NOT NULL
  )
"@
            } catch {
                $sheetIndexCompletionData = 0L
                $sheetPackageCompletionData = 0L
            }
        }

        $summary.database = @{
            rowsMatched = $dbCounts
            deleted = [ordered]@{}
            sheetIndexMatched = $sheetIndexMatched
            sheetIndexLaneMatched = $sheetIndexLaneMatched
            sheetIndexNonLaneMatched = $sheetIndexNonLaneMatched
            sheetDocumentsQcPdfMatched = $sheetDocumentsQcPdfMatched
            sheetIndexCompletionData = $sheetIndexCompletionData
            sheetPackagesMatched = $sheetPackagesMatched
            sheetDocumentsMatched = $sheetDocumentsMatched
            sheetPackageCompletionData = $sheetPackageCompletionData
            sheetPackageQcPdfsMatched = $sheetPackageQcPdfsMatched
            queryTimeoutSeconds = $QueryTimeoutSeconds
        }
        if (-not $SkipPreviewCounts.IsPresent) {
            Write-Host '  Telemetry rows to delete:' -ForegroundColor Gray
            foreach ($k in @($dbCounts.Keys)) {
                Write-Host ("    {0}: {1}" -f $k, $dbCounts[$k]) -ForegroundColor Gray
            }
        }
        if ($SkipSheetIndexUpdate.IsPresent) {
            Write-Host '  sheet_index / sheet_packages / sheet_documents: skipped (-SkipSheetIndexUpdate); no rows deleted or updated.' -ForegroundColor DarkGray
        } else {
            if ($KeepLanePdfRegistry.IsPresent) {
                Write-Host ('  sheet_index: {0} row(s) matched - UPDATE only, rows are not deleted.' -f $sheetIndexMatched) -ForegroundColor DarkGray
            } else {
                Write-Host ('  sheet_index: {0} row(s) matched - {1} lane PDF row(s) DELETE, {2} stem/DGN row(s) UPDATE.' -f `
                    $sheetIndexMatched, $sheetIndexLaneMatched, $sheetIndexNonLaneMatched) -ForegroundColor DarkGray
                Write-Host ('  sheet_documents qc_pdf role: {0} row(s) matched - DELETE on confirm.' -f $sheetDocumentsQcPdfMatched) -ForegroundColor DarkGray
            }
            Write-Host ('  sheet_index completion/cycle reset: {0} row(s) with qc_cycle_id or completion counts/timestamps to clear.' -f $sheetIndexCompletionData) -ForegroundColor DarkGray
            Write-Host ('  sheet_packages: {0} row(s) matched - UPDATE only, rows are not deleted.' -f $sheetPackagesMatched) -ForegroundColor DarkGray
            Write-Host ('  sheet_packages completion/cycle reset: {0} row(s) with qc_cycle_id or completion counts/timestamps to clear.' -f $sheetPackageCompletionData) -ForegroundColor DarkGray
            Write-Host ('  sheet_documents: {0} row(s) matched - pw_state_name UPDATE only, rows are not deleted.' -f $sheetDocumentsMatched) -ForegroundColor DarkGray
            Write-Host ('  sheet_package_qc_pdfs: {0} row(s) matched - DELETE on confirm.' -f $sheetPackageQcPdfsMatched) -ForegroundColor DarkGray
        }

        if ($script:_rqcfInteractiveLaunch -and -not $doApply) {
            Write-Host ''
            $confirmAnswer = (Read-Host 'Apply database reset? (y/N)').Trim()
            if ($confirmAnswer -match '^[yY]') {
                $doApply = $true
                $DryRun = $false
                $summary.dryRun = $false
            }
        }

        if ($doApply) {
            Write-Host '[Database] Deleting telemetry rows...' -ForegroundColor Cyan
            $deleted = $summary.database.deleted

            if ($IncludeCommentTelemetry.IsPresent) {
                $whereCommentHistory = _RQCF-TelemetryScopeClause -DocumentGuidColumn 'h.document_id' @scopeArgs
                $deleted.qc_comment_status_history = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_comment_status_history' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_comment_status_history
WHERE history_id IN (
  SELECT h.history_id FROM qc_comment_status_history h
  WHERE ($whereCommentHistory) OR h.detected_run_id IN ($commentRunSub)
)
"@
                $deleted.qc_comments = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_comments' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_comments
WHERE comment_record_id IN (SELECT c.comment_record_id FROM qc_comments c WHERE c.run_id IN ($commentRunSub))
"@
                $deleted.qc_comment_runs = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_comment_runs' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_comment_runs
WHERE run_id IN ($commentRunSub)
"@
            }

            try {
                $deleted.qc_notification_messages = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_notification_messages' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_notification_messages
WHERE id IN (
  SELECT m.id FROM qc_notification_messages m
  WHERE m.notification_log_id IN (SELECT n.id FROM notification_log n WHERE ($whereNotification))
)
"@
            } catch {
                Write-Host ("  [qc_notification_messages] skipped: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
            }

            $deleted.notification_log = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'notification_log' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM notification_log
WHERE id IN (SELECT n.id FROM notification_log n WHERE ($whereNotification))
"@
            $deleted.qc_workflow_events = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_workflow_events' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_workflow_events
WHERE event_id IN (SELECT w.event_id FROM qc_workflow_events w WHERE ($whereWorkflow))
"@
            $deleted.qc_cycle_completions = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_cycle_completions' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_cycle_completions
WHERE id IN (SELECT c.id FROM qc_cycle_completions c WHERE ($whereCompletion))
"@
            $deleted.transition_events = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'transition_events' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM transition_events
WHERE id IN (SELECT t.id FROM transition_events t WHERE ($whereTransition))
"@
            $deleted.document_state_history = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'document_state_history' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM document_state_history
WHERE id IN (SELECT h.id FROM document_state_history h WHERE ($whereHistory))
"@
            $deleted.document_activity = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'document_activity' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM document_activity
WHERE id IN (SELECT d.id FROM document_activity d WHERE ($whereActivity))
"@
            $deleted.audit_events = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'audit_events' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM audit_events
WHERE id IN (SELECT a.id FROM audit_events a WHERE ($whereAudit))
"@
            $deleted.processing_jobs = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'processing_jobs' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM processing_jobs
WHERE id IN (SELECT j.id FROM processing_jobs j WHERE ($whereJobs))
"@
            try {
                $deleted.sheet_package_qc_pdfs = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'sheet_package_qc_pdfs' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM sheet_package_qc_pdfs
WHERE id IN (
  SELECT q.id FROM sheet_package_qc_pdfs q
  WHERE EXISTS (SELECT 1 FROM sheet_packages sp WHERE sp.sheet_package_id = q.sheet_package_id AND ($spFolderClause))
)
"@
            } catch {
                Write-Host ("  [sheet_package_qc_pdfs] skipped: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
            }

            foreach ($k in @($deleted.Keys)) {
                $summary.totalRowsDeleted += [long]$deleted[$k]
            }
            if ($null -ne $deleted['sheet_package_qc_pdfs']) {
                $summary.sheetPackageQcPdfsDeleted = [long]$deleted['sheet_package_qc_pdfs']
            }

            if (-not $SkipSheetIndexUpdate.IsPresent) {
                $purgeLaneRegistry = (-not $KeepLanePdfRegistry.IsPresent) -or $scopeIsQcPdfOnly
                if ($purgeLaneRegistry) {
                    try {
                        $deleted.sheet_index_lane = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'sheet_index_lane_pdfs' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM sheet_index
WHERE document_guid IN (
  SELECT si.document_guid FROM sheet_index si
  WHERE ($siFolderClause)$(if (-not $scopeIsQcPdfOnly) { " AND ($siLaneNameClause)" } else { '' })
)
"@
                        $summary.sheetIndexLaneRowsDeleted = [long]$deleted['sheet_index_lane']
                    } catch {
                        Write-Host ("  [sheet_index lane PDF rows] skipped: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
                    }
                    try {
                        $sdQcPdfWhere = if ($scopeIsQcPdfOnly) {
                            "($spFolderClause) AND sd.document_role = 'qc_pdf' AND LOWER(LTRIM(RTRIM(sd.document_name))) = LOWER(LTRIM(RTRIM(@scopeDocName)))"
                        } else {
                            "($spFolderClause) AND sd.document_role = 'qc_pdf'"
                        }
                        $deleted.sheet_documents_qc_pdf = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'sheet_documents_qc_pdf' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM sheet_documents
WHERE document_guid IN (
  SELECT sd.document_guid FROM sheet_documents sd
  INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
  WHERE $sdQcPdfWhere
)
"@
                        $summary.sheetDocumentsQcPdfDeleted = [long]$deleted['sheet_documents_qc_pdf']
                    } catch {
                        Write-Host ("  [sheet_documents qc_pdf] skipped: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
                    }
                }

                if ($scopeIsQcPdfOnly) {
                    $laneType = _RQCF-GetLaneTypeFromDocumentName -DocumentName $scopeDocName
                    $packageLaneClearSql = switch ($laneType) {
                        'production' { "qc_pdf_guid = NULL, qc_pdf_name = NULL," }
                        'check' { "qc_chk_pdf_guid = NULL, qc_chk_pdf_name = NULL," }
                        'review' { "qc_rev_pdf_guid = NULL, qc_rev_pdf_name = NULL," }
                        default { '' }
                    }
                    if ($packageLaneClearSql) {
                        $packageSql = @"
UPDATE sp
SET last_updated_at = SYSDATETIMEOFFSET(),
$packageLaneClearSql
    qc_process_type = NULL
FROM sheet_packages sp
WHERE ($spFolderClause)
"@
                        if ($PSCmdlet.ShouldProcess('sheet_packages', 'CLEAR lane QC PDF pairing')) {
                            try {
                                $summary.sheetPackagesUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $packageSql -Params $params -CommandTimeout $QueryTimeoutSeconds
                                Write-Host ('  [sheet_packages] cleared {0} lane pairing on {1} row(s).' -f $laneType, $summary.sheetPackagesUpdated) -ForegroundColor Green
                            } catch {
                                Write-Host ("  [sheet_packages] lane clear failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                            }
                        }
                    }
                    if ($summary.sheetIndexLaneRowsDeleted -gt 0) {
                        Write-Host ('  [sheet_index] deleted {0} lane PDF row(s).' -f $summary.sheetIndexLaneRowsDeleted) -ForegroundColor Green
                    }
                    if ($summary.sheetDocumentsQcPdfDeleted -gt 0) {
                        Write-Host ('  [sheet_documents] deleted {0} qc_pdf role row(s).' -f $summary.sheetDocumentsQcPdfDeleted) -ForegroundColor Green
                    }
                } else {

                $completionResetSql = @"
    production_qc_completed_count = 0,
    production_qc_last_completed_at = NULL,
    peer_review_completed_count = 0,
    peer_review_last_completed_at = NULL,
    independent_check_completed_count = 0,
    independent_check_last_completed_at = NULL,
"@
                $sheetIndexLaneResetSql = @"
    qc_process_type = NULL,
    qc_pdf_guid = NULL,
    qc_pdf_name = NULL,
"@
                $sheetIndexQcFieldResetSql = ''
                if (-not $KeepSheetIndexQcFields.IsPresent) {
                    $sheetIndexQcFieldResetSql = @"
    qc_review_type = NULL,
    qc_assigned_to = NULL,
"@
                }
                if (-not $KeepSheetIndexQcFields.IsPresent) {
                    $sheetSql = @"
UPDATE si
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
    qc_stage = NULL,
    qc_status = NULL,
    last_audit_event_at = NULL,
$completionResetSql
$sheetIndexLaneResetSql
$sheetIndexQcFieldResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL
FROM sheet_index si
WHERE ($siFolderClause)$(if (-not $KeepLanePdfRegistry.IsPresent) { " AND NOT ($siLaneNameClause)" } else { '' })
"@
                } else {
                    $sheetSql = @"
UPDATE si
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
$completionResetSql
$sheetIndexLaneResetSql
$sheetIndexQcFieldResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL
FROM sheet_index si
WHERE ($siFolderClause)$(if (-not $KeepLanePdfRegistry.IsPresent) { " AND NOT ($siLaneNameClause)" } else { '' })
"@
                }

                if ($PSCmdlet.ShouldProcess('sheet_index', 'UPDATE pw_state_name')) {
                    try {
                        $summary.sheetIndexUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $sheetSql -Params $params -CommandTimeout $QueryTimeoutSeconds
                        Write-Host ('  [sheet_index] updated {0} row(s): pw_state_name, qc_process_type, lane QC pairing, qc_cycle/completion fields cleared (not deleted).' -f $summary.sheetIndexUpdated) -ForegroundColor Green
                        if (-not $KeepLanePdfRegistry.IsPresent -and $summary.sheetIndexLaneRowsDeleted -gt 0) {
                            Write-Host ('  [sheet_index] deleted {0} lane PDF row(s) (*-prod/-chk/-rev).' -f $summary.sheetIndexLaneRowsDeleted) -ForegroundColor Green
                        }
                        if (-not $KeepLanePdfRegistry.IsPresent -and $summary.sheetDocumentsQcPdfDeleted -gt 0) {
                            Write-Host ('  [sheet_documents] deleted {0} qc_pdf role row(s).' -f $summary.sheetDocumentsQcPdfDeleted) -ForegroundColor Green
                        }
                    } catch {
                        Write-Host ("  [sheet_index] extended reset failed ({0}); retrying without v1.18 columns." -f $_.Exception.Message) -ForegroundColor Yellow
                        $sheetSqlLegacy = if (-not $KeepSheetIndexQcFields.IsPresent) {
                            @"
UPDATE si
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
    qc_stage = NULL,
    qc_status = NULL,
    last_audit_event_at = NULL,
$completionResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL
FROM sheet_index si
WHERE ($siFolderClause)$(if (-not $KeepLanePdfRegistry.IsPresent) { " AND NOT ($siLaneNameClause)" } else { '' })
"@
                        } else {
                            @"
UPDATE si
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
$completionResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL
FROM sheet_index si
WHERE ($siFolderClause)$(if (-not $KeepLanePdfRegistry.IsPresent) { " AND NOT ($siLaneNameClause)" } else { '' })
"@
                        }
                        $summary.sheetIndexUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $sheetSqlLegacy -Params $params -CommandTimeout $QueryTimeoutSeconds
                        Write-Host ('  [sheet_index] updated {0} row(s) (legacy columns only).' -f $summary.sheetIndexUpdated) -ForegroundColor Green
                    }
                }

                $packageResetSql = @"
    production_qc_completed_count = 0,
    production_qc_last_completed_at = NULL,
    peer_review_completed_count = 0,
    peer_review_last_completed_at = NULL,
    independent_check_completed_count = 0,
    independent_check_last_completed_at = NULL,
"@
                $packageSql = @"
UPDATE sp
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
$packageResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL,
    qc_review_type = NULL,
    qc_assigned_to = NULL,
    qc_process_type = NULL,
    qc_pdf_guid = NULL,
    qc_pdf_name = NULL,
    qc_chk_pdf_guid = NULL,
    qc_chk_pdf_name = NULL,
    qc_rev_pdf_guid = NULL,
    qc_rev_pdf_name = NULL
FROM sheet_packages sp
WHERE ($spFolderClause)
"@
                if ($PSCmdlet.ShouldProcess('sheet_packages', 'UPDATE pw_state_name')) {
                    try {
                        $summary.sheetPackagesUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $packageSql -Params $params -CommandTimeout $QueryTimeoutSeconds
                        Write-Host ('  [sheet_packages] updated {0} row(s): pw_state_name, qc_process_type, lane PDF aliases, qc_cycle/completion fields cleared (not deleted).' -f $summary.sheetPackagesUpdated) -ForegroundColor Green
                    } catch {
                        Write-Host ("  [sheet_packages] extended reset failed ({0}); retrying without v1.18 lane columns." -f $_.Exception.Message) -ForegroundColor Yellow
                        $packageSqlLegacy = @"
UPDATE sp
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
$packageResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL,
    qc_review_type = NULL,
    qc_assigned_to = NULL
FROM sheet_packages sp
WHERE ($spFolderClause)
"@
                        $summary.sheetPackagesUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $packageSqlLegacy -Params $params -CommandTimeout $QueryTimeoutSeconds
                        Write-Host ('  [sheet_packages] updated {0} row(s) (legacy columns only).' -f $summary.sheetPackagesUpdated) -ForegroundColor Green
                    }
                }

                $docSql = @"
UPDATE sd
SET pw_state_name = @targetState,
    last_seen_at = SYSDATETIMEOFFSET()
FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE ($spFolderClause)
"@
                if ($PSCmdlet.ShouldProcess('sheet_documents', 'UPDATE pw_state_name')) {
                    $summary.sheetDocumentsUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $docSql -Params $params -CommandTimeout $QueryTimeoutSeconds
                    Write-Host ('  [sheet_documents] updated {0} row(s): pw_state_name reset (not deleted).' -f $summary.sheetDocumentsUpdated) -ForegroundColor Green
                }
                }

            } else {
                Write-Host '  [sheet_index / sheet_packages / sheet_documents] skipped (-SkipSheetIndexUpdate).' -ForegroundColor DarkGray
            }

            Write-Host ("Done. Deleted {0} telemetry row(s)." -f $summary.totalRowsDeleted) -ForegroundColor Green
        } else {
            if ($script:_rqcfInteractiveLaunch) {
                Write-Host 'Database reset cancelled.' -ForegroundColor Yellow
            } else {
                Write-Host 'Dry run: no database rows deleted. Pass -ConfirmReset to apply.' -ForegroundColor Yellow
            }
        }
    } finally {
        try { $sessionRes.Data.session.Dispose() } catch { }
    }
} else {
    Write-Host '[Database] Skipped (-SkipDatabase).' -ForegroundColor DarkGray
    $summary.database = @{ skipped = $true }
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }

if ($script:_rqcfInteractiveLaunch) {
    _RQCF-PauseIfInteractiveConsole
}
