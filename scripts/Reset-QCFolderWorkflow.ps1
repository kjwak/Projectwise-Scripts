<#
.SYNOPSIS
Resets ProjectWise workflow state for all documents in a folder and clears QC telemetry in SQL.

.DESCRIPTION
For a single Sheets (or any) folder:

1. ProjectWise - sets workflow state on every document in the folder to the target state
   (default: qcWorkflow.states.production, usually "In Production").
2. Database - deletes folder-scoped rows from QC telemetry tables (not sheet_index, sheet_packages, or sheet_documents).

Default is preview only. Pass -ConfirmReset to apply PW and database changes.

sheet_index, sheet_packages, and sheet_documents rows are never deleted. By default the script updates pw_state_name (and clears
qc_stage/qc_status on sheet_index unless -KeepSheetIndexQcFields), clears qc_cycle_id/qc_cycle_number, zeros
production_qc_completed_count, production_qc_last_completed_at, peer_review_completed_count,
peer_review_last_completed_at, independent_check_completed_count, and
independent_check_last_completed_at on sheet_index and sheet_packages, clears qc_review_type/qc_assigned_to on
sheet_packages, updates sheet_documents.pw_state_name, and deletes qc_cycle_completions rows for the folder
(by document_guid or sheet_package_id).Pass -SkipSheetIndexUpdate to leave sheet_index completely unchanged. Does not remove queue JSON jobs
(use Purge-QCPendingByFilters or manual queue cleanup separately).

Does not delete ProjectWise documents, sheet_index rows, sheet_packages rows, or sheet_documents rows.

.EXAMPLE
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath 'AZFWY1704-FD02-SR202\CADD\Sheets' -ConfirmReset
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipProjectWise
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipDatabase
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipSheetIndexUpdate
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$FolderPath,

    [string]$AppSettingsPath = '',
    [string]$TargetState = '',
    [switch]$ConfirmReset,
    [switch]$DryRun,
    [switch]$SkipProjectWise,
    [switch]$SkipDatabase,
    [switch]$IncludeCommentTelemetry,
    [switch]$KeepSheetIndexQcFields,
    [switch]$SkipSheetIndexUpdate,
    [int]$BatchSize = 5000,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot 'Import-QCScriptModules.ps1') -RepoRoot $repoRoot -AdditionalModules @(
    'Core.Paths.psm1'
    'PW.Connection.psm1'
    'PW.Discovery.psm1'
)

foreach ($moduleName in @('pwps', 'pwps_dab')) {
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
    Write-Host 'Preview only: pass -ConfirmReset to apply PW + database reset, or -DryRun explicitly.' -ForegroundColor Yellow
    $DryRun = $true
}

$normRes = Normalize-QCDocumentsFolderPath -Path $FolderPath
if (-not $normRes.IsSuccess) { throw $normRes.Message }
$normFolder = [string]$normRes.Data.path

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
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

function _RQCF-SheetIndexGuidSubquery {
    param([string]$FolderLikeClause)
    return @"
SELECT LTRIM(RTRIM(si.document_guid))
FROM sheet_index si
WHERE si.document_guid IS NOT NULL AND ($FolderLikeClause)
UNION
SELECT LTRIM(RTRIM(si.qc_pdf_guid))
FROM sheet_index si
WHERE si.qc_pdf_guid IS NOT NULL AND ($FolderLikeClause)
"@
}

function _RQCF-SheetPackageIdSubquery {
    param([string]$FolderLikeClause)
    return @"
SELECT sp.sheet_package_id
FROM sheet_packages sp
WHERE ($FolderLikeClause)
"@
}

function _RQCF-FolderDocumentGuidSubquery {
    param(
        [string]$SheetIndexFolderClause,
        [string]$SheetPackageFolderClause
    )
    return @"
SELECT LTRIM(RTRIM(si.document_guid))
FROM sheet_index si
WHERE si.document_guid IS NOT NULL AND ($SheetIndexFolderClause)
UNION
SELECT LTRIM(RTRIM(si.qc_pdf_guid))
FROM sheet_index si
WHERE si.qc_pdf_guid IS NOT NULL AND ($SheetIndexFolderClause)
UNION
SELECT LTRIM(RTRIM(CAST(sp.dgn_guid AS NVARCHAR(50))))
FROM sheet_packages sp
WHERE sp.dgn_guid IS NOT NULL AND ($SheetPackageFolderClause)
UNION
SELECT LTRIM(RTRIM(CAST(sp.sheet_pdf_guid AS NVARCHAR(50))))
FROM sheet_packages sp
WHERE sp.sheet_pdf_guid IS NOT NULL AND ($SheetPackageFolderClause)
UNION
SELECT LTRIM(RTRIM(CAST(sp.qc_pdf_guid AS NVARCHAR(50))))
FROM sheet_packages sp
WHERE sp.qc_pdf_guid IS NOT NULL AND ($SheetPackageFolderClause)
UNION
SELECT LTRIM(RTRIM(CAST(sd.document_guid AS NVARCHAR(50))))
FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE sd.document_guid IS NOT NULL AND ($SheetPackageFolderClause)
"@
}

function _RQCF-GetScalar {
    param([hashtable]$Config, [string]$Sql, [hashtable]$Params)
    $res = Invoke-QCDatabaseScalar -Config $Config -Sql $Sql -Parameters $Params
    if (-not $res.IsSuccess) { throw $res.Message }
    return [long]$res.Data.value
}

function _RQCF-RunDeleteLoop {
    param(
        [hashtable]$Config,
        [string]$Sql,
        [hashtable]$Params,
        [string]$Label
    )
    $total = 0L
    do {
        $batch = 0
        if ($PSCmdlet.ShouldProcess($Label, 'DELETE batch')) {
            $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $Sql -Parameters $Params
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
    folderLike  = _RQCF-NormalizePathPattern -Pattern $normFolder
}
$siFolderClause = _RQCF-PathLikeClause -ColumnSql 'si.folder_path' -Params $params -ParamName 'folderLike'
$spFolderClause = _RQCF-PathLikeClause -ColumnSql 'sp.folder_path' -Params $params -ParamName 'folderLike'
$folderPathClause = _RQCF-PathLikeClause -ColumnSql 'folder_path' -Params $params -ParamName 'folderLike'
$resolvedFolderClause = _RQCF-PathLikeClause -ColumnSql 'resolved_folder' -Params $params -ParamName 'folderLike'
$sourceFolderClause = _RQCF-PathLikeClause -ColumnSql 'source_folder' -Params $params -ParamName 'folderLike'
$guidSub = _RQCF-FolderDocumentGuidSubquery -SheetIndexFolderClause $siFolderClause -SheetPackageFolderClause $spFolderClause
$packageIdSub = _RQCF-SheetPackageIdSubquery -FolderLikeClause $spFolderClause

$summary = [ordered]@{
    dryRun            = $DryRun.IsPresent
    folderPath        = $normFolder
    targetState       = $TargetState
    projectWise       = $null
    database          = $null
    totalRowsDeleted  = 0
    sheetIndexUpdated = 0
    sheetPackagesUpdated = 0
    sheetDocumentsUpdated = 0
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
    Write-Host '[Database] Counting folder-scoped telemetry rows...' -ForegroundColor Cyan

    $dbCounts = [ordered]@{}
    $dbCounts.notification_log = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM notification_log
WHERE document_guid IN ($guidSub) OR ($folderPathClause) OR sheet_package_id IN ($packageIdSub)
"@
    $dbCounts.document_state_history = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM document_state_history
WHERE document_guid IN ($guidSub) OR ($folderPathClause) OR sheet_package_id IN ($packageIdSub)
"@
    $dbCounts.transition_events = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM transition_events
WHERE document_guid IN ($guidSub) OR ($folderPathClause) OR sheet_package_id IN ($packageIdSub)
"@
    $dbCounts.document_activity = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM document_activity
WHERE document_guid IN ($guidSub) OR ($folderPathClause)
"@
    $dbCounts.qc_workflow_events = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM qc_workflow_events w
WHERE w.document_id IN ($guidSub) OR w.sheet_package_id IN ($packageIdSub)
"@
    $dbCounts.qc_cycle_completions = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM qc_cycle_completions
WHERE document_guid IN ($guidSub) OR sheet_package_id IN ($packageIdSub)
"@
    $dbCounts.audit_events = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM audit_events
WHERE ($resolvedFolderClause) OR pw_objguid IN ($guidSub)
"@
    $dbCounts.processing_jobs = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM processing_jobs
WHERE ($sourceFolderClause) OR sheet_package_id IN ($packageIdSub)
"@

    $sheetPackagesMatched = 0L
    $sheetDocumentsMatched = 0L
    $sheetPackageCompletionData = 0L
    try {
        $sheetPackagesMatched = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_packages sp WHERE ($spFolderClause)
"@
        $sheetDocumentsMatched = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE ($spFolderClause)
"@
    } catch {
        $sheetPackagesMatched = 0L
        $sheetDocumentsMatched = 0L
    }

    $sheetIndexMatched = 0L
    $sheetIndexCompletionData = 0L
    if (-not $SkipSheetIndexUpdate.IsPresent) {
        $sheetIndexMatched = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si WHERE ($siFolderClause)
"@
        try {
            $sheetIndexCompletionData = _RQCF-GetScalar -Config $config -Params $params -Sql @"
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
        } catch {
            $sheetIndexCompletionData = 0L
        }
        try {
            $sheetPackageCompletionData = _RQCF-GetScalar -Config $config -Params $params -Sql @"
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
            $sheetPackageCompletionData = 0L
        }
    }

    if ($IncludeCommentTelemetry.IsPresent) {
        $runSub = @"
SELECT r.run_id FROM qc_comment_runs r
WHERE r.document_id IN ($guidSub)
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @folderLike
"@
        $dbCounts.qc_comment_status_history = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM qc_comment_status_history h
WHERE h.document_id IN ($guidSub) OR h.detected_run_id IN ($runSub)
"@
        $dbCounts.qc_comments = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM qc_comments c WHERE c.run_id IN ($runSub)
"@
        $dbCounts.qc_comment_runs = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM qc_comment_runs r
WHERE r.document_id IN ($guidSub)
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @folderLike
"@
    }

    $summary.database = @{
        rowsMatched = $dbCounts
        deleted = [ordered]@{}
        sheetIndexMatched = $sheetIndexMatched
        sheetIndexCompletionData = $sheetIndexCompletionData
        sheetPackagesMatched = $sheetPackagesMatched
        sheetDocumentsMatched = $sheetDocumentsMatched
        sheetPackageCompletionData = $sheetPackageCompletionData
    }
    Write-Host '  Telemetry rows to delete:' -ForegroundColor Gray
    foreach ($k in @($dbCounts.Keys)) {
        Write-Host ("    {0}: {1}" -f $k, $dbCounts[$k]) -ForegroundColor Gray
    }
    if ($SkipSheetIndexUpdate.IsPresent) {
        Write-Host '  sheet_index / sheet_packages / sheet_documents: skipped (-SkipSheetIndexUpdate); no rows deleted or updated.' -ForegroundColor DarkGray
    } else {
        Write-Host ('  sheet_index: {0} row(s) matched - UPDATE only, rows are not deleted.' -f $sheetIndexMatched) -ForegroundColor DarkGray
        Write-Host ('  sheet_index completion/cycle reset: {0} row(s) with qc_cycle_id or completion counts/timestamps to clear.' -f $sheetIndexCompletionData) -ForegroundColor DarkGray
        Write-Host ('  sheet_packages: {0} row(s) matched - UPDATE only, rows are not deleted.' -f $sheetPackagesMatched) -ForegroundColor DarkGray
        Write-Host ('  sheet_packages completion/cycle reset: {0} row(s) with qc_cycle_id or completion counts/timestamps to clear.' -f $sheetPackageCompletionData) -ForegroundColor DarkGray
        Write-Host ('  sheet_documents: {0} row(s) matched - pw_state_name UPDATE only, rows are not deleted.' -f $sheetDocumentsMatched) -ForegroundColor DarkGray
    }

    if ($doApply) {
        Write-Host '[Database] Deleting telemetry rows...' -ForegroundColor Cyan
        $deleted = $summary.database.deleted

        if ($IncludeCommentTelemetry.IsPresent) {
            $runSub = @"
SELECT r.run_id FROM qc_comment_runs r
WHERE r.document_id IN ($guidSub)
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @folderLike
"@
            $deleted.qc_comment_status_history = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'qc_comment_status_history' -Sql @"
DELETE TOP (@batchSize) FROM qc_comment_status_history
WHERE history_id IN (
  SELECT h.history_id FROM qc_comment_status_history h
  WHERE h.document_id IN ($guidSub) OR h.detected_run_id IN ($runSub)
)
"@
            $deleted.qc_comments = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'qc_comments' -Sql @"
DELETE TOP (@batchSize) FROM qc_comments
WHERE comment_record_id IN (SELECT c.comment_record_id FROM qc_comments c WHERE c.run_id IN ($runSub))
"@
            $deleted.qc_comment_runs = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'qc_comment_runs' -Sql @"
DELETE TOP (@batchSize) FROM qc_comment_runs
WHERE run_id IN ($runSub)
"@
        }

        $deleted.notification_log = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'notification_log' -Sql @"
DELETE TOP (@batchSize) FROM notification_log
WHERE id IN (
  SELECT n.id FROM notification_log n
  WHERE n.document_guid IN ($guidSub) OR ($folderPathClause) OR n.sheet_package_id IN ($packageIdSub)
)
"@
        $deleted.qc_workflow_events = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'qc_workflow_events' -Sql @"
DELETE TOP (@batchSize) FROM qc_workflow_events
WHERE event_id IN (
  SELECT w.event_id FROM qc_workflow_events w
  WHERE w.document_id IN ($guidSub) OR w.sheet_package_id IN ($packageIdSub)
)
"@
        $deleted.qc_cycle_completions = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'qc_cycle_completions' -Sql @"
DELETE TOP (@batchSize) FROM qc_cycle_completions
WHERE id IN (
  SELECT c.id FROM qc_cycle_completions c
  WHERE c.document_guid IN ($guidSub) OR c.sheet_package_id IN ($packageIdSub)
)
"@
        $deleted.transition_events = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'transition_events' -Sql @"
DELETE TOP (@batchSize) FROM transition_events
WHERE id IN (
  SELECT t.id FROM transition_events t
  WHERE t.document_guid IN ($guidSub) OR ($folderPathClause) OR t.sheet_package_id IN ($packageIdSub)
)
"@
        $deleted.document_state_history = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'document_state_history' -Sql @"
DELETE TOP (@batchSize) FROM document_state_history
WHERE id IN (
  SELECT h.id FROM document_state_history h
  WHERE h.document_guid IN ($guidSub) OR ($folderPathClause) OR h.sheet_package_id IN ($packageIdSub)
)
"@
        $deleted.document_activity = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'document_activity' -Sql @"
DELETE TOP (@batchSize) FROM document_activity
WHERE id IN (
  SELECT d.id FROM document_activity d
  WHERE d.document_guid IN ($guidSub) OR ($folderPathClause)
)
"@
        $deleted.audit_events = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'audit_events' -Sql @"
DELETE TOP (@batchSize) FROM audit_events
WHERE id IN (
  SELECT a.id FROM audit_events a
  WHERE ($resolvedFolderClause) OR a.pw_objguid IN ($guidSub)
)
"@
        $deleted.processing_jobs = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'processing_jobs' -Sql @"
DELETE TOP (@batchSize) FROM processing_jobs
WHERE id IN (
  SELECT j.id FROM processing_jobs j
  WHERE ($sourceFolderClause) OR j.sheet_package_id IN ($packageIdSub)
)
"@

        foreach ($k in @($deleted.Keys)) {
            $summary.totalRowsDeleted += [long]$deleted[$k]
        }

        if (-not $SkipSheetIndexUpdate.IsPresent) {
            $completionResetSql = @"
    production_qc_completed_count = 0,
    production_qc_last_completed_at = NULL,
    peer_review_completed_count = 0,
    peer_review_last_completed_at = NULL,
    independent_check_completed_count = 0,
    independent_check_last_completed_at = NULL,
"@
            if (-not $KeepSheetIndexQcFields.IsPresent) {
                $sheetSql = @"
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
WHERE ($siFolderClause)
"@
            } else {
                $sheetSql = @"
UPDATE si
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET(),
$completionResetSql
    qc_cycle_id = NULL,
    qc_cycle_number = NULL
FROM sheet_index si
WHERE ($siFolderClause)
"@
            }

            if ($PSCmdlet.ShouldProcess('sheet_index', 'UPDATE pw_state_name')) {
                $upd = Invoke-QCDatabaseNonQuery -Config $config -Sql $sheetSql -Parameters $params
                if (-not $upd.IsSuccess) { throw $upd.Message }
                $summary.sheetIndexUpdated = [int]$upd.Data.rowsAffected
                Write-Host ('  [sheet_index] updated {0} row(s): pw_state_name, qc_cycle_id/number, production/peer/independent completion counts and last_completed_at cleared (not deleted).' -f $summary.sheetIndexUpdated) -ForegroundColor Green
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
    qc_assigned_to = NULL
FROM sheet_packages sp
WHERE ($spFolderClause)
"@
            if ($PSCmdlet.ShouldProcess('sheet_packages', 'UPDATE pw_state_name')) {
                $pkgUpd = Invoke-QCDatabaseNonQuery -Config $config -Sql $packageSql -Parameters $params
                if (-not $pkgUpd.IsSuccess) { throw $pkgUpd.Message }
                $summary.sheetPackagesUpdated = [int]$pkgUpd.Data.rowsAffected
                Write-Host ('  [sheet_packages] updated {0} row(s): pw_state_name, qc_cycle_id/number, review fields, completion counts/timestamps cleared (not deleted).' -f $summary.sheetPackagesUpdated) -ForegroundColor Green
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
                $docUpd = Invoke-QCDatabaseNonQuery -Config $config -Sql $docSql -Parameters $params
                if (-not $docUpd.IsSuccess) { throw $docUpd.Message }
                $summary.sheetDocumentsUpdated = [int]$docUpd.Data.rowsAffected
                Write-Host ('  [sheet_documents] updated {0} row(s): pw_state_name reset (not deleted).' -f $summary.sheetDocumentsUpdated) -ForegroundColor Green
            }
        } else {
            Write-Host '  [sheet_index / sheet_packages / sheet_documents] skipped (-SkipSheetIndexUpdate).' -ForegroundColor DarkGray
        }

        Write-Host ("Done. Deleted {0} telemetry row(s)." -f $summary.totalRowsDeleted) -ForegroundColor Green
    } else {
        Write-Host 'Dry run: no database rows deleted. Pass -ConfirmReset to apply.' -ForegroundColor Yellow
    }
} else {
    Write-Host '[Database] Skipped (-SkipDatabase).' -ForegroundColor DarkGray
    $summary.database = @{ skipped = $true }
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }
