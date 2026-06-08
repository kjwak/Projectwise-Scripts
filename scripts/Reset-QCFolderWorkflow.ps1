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
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipPreviewCounts -QueryTimeoutSeconds 600
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
    [int]$QueryTimeoutSeconds = 300,
    [switch]$SkipPreviewCounts,
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

function _RQCF-GetScopeInitSql {
    param(
        [string]$SheetIndexFolderClause,
        [string]$SheetPackageFolderClause
    )
    return @"
IF OBJECT_ID('tempdb..#rqcf_pkg') IS NOT NULL DROP TABLE #rqcf_pkg;
CREATE TABLE #rqcf_pkg (sheet_package_id UNIQUEIDENTIFIER NOT NULL PRIMARY KEY);
INSERT INTO #rqcf_pkg (sheet_package_id)
SELECT sp.sheet_package_id
FROM sheet_packages sp
WHERE ($SheetPackageFolderClause);

IF OBJECT_ID('tempdb..#rqcf_doc') IS NOT NULL DROP TABLE #rqcf_doc;
CREATE TABLE #rqcf_doc (document_guid NVARCHAR(50) NOT NULL PRIMARY KEY);
INSERT INTO #rqcf_doc (document_guid)
SELECT DISTINCT v.document_guid
FROM (
    SELECT LTRIM(RTRIM(CAST(sd.document_guid AS NVARCHAR(50)))) AS document_guid
    FROM sheet_documents sd
    INNER JOIN #rqcf_pkg p ON p.sheet_package_id = sd.sheet_package_id
    WHERE sd.document_guid IS NOT NULL
    UNION ALL
    SELECT LTRIM(RTRIM(si.document_guid))
    FROM sheet_index si
    WHERE si.document_guid IS NOT NULL AND ($SheetIndexFolderClause)
    UNION ALL
    SELECT LTRIM(RTRIM(si.qc_pdf_guid))
    FROM sheet_index si
    WHERE si.qc_pdf_guid IS NOT NULL AND ($SheetIndexFolderClause)
) v
WHERE v.document_guid <> '';
"@
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
$scopeInitSql = _RQCF-GetScopeInitSql -SheetIndexFolderClause $siFolderClause -SheetPackageFolderClause $spFolderClause
$nlFolderClause = _RQCF-PathLikeClause -ColumnSql 'n.folder_path' -Params $params -ParamName 'folderLike'
$histFolderClause = _RQCF-PathLikeClause -ColumnSql 'h.folder_path' -Params $params -ParamName 'folderLike'
$trFolderClause = _RQCF-PathLikeClause -ColumnSql 't.folder_path' -Params $params -ParamName 'folderLike'
$actFolderClause = _RQCF-PathLikeClause -ColumnSql 'd.folder_path' -Params $params -ParamName 'folderLike'
$auditFolderClause = _RQCF-PathLikeClause -ColumnSql 'a.resolved_folder' -Params $params -ParamName 'folderLike'
$jobFolderClause = _RQCF-PathLikeClause -ColumnSql 'j.source_folder' -Params $params -ParamName 'folderLike'

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
    $sessionRes = New-QCDatabaseSession -Config $config
    if (-not $sessionRes.IsSuccess) { throw $sessionRes.Message }
    $dbConn = $sessionRes.Data.session.connection
    try {
        Write-Host '[Database] Building folder scope (#rqcf_pkg / #rqcf_doc)...' -ForegroundColor Cyan
        [void]_RQCF-RunNonQueryConn -Connection $dbConn -Sql $scopeInitSql -Params $params -CommandTimeout $QueryTimeoutSeconds

        $scopePkgIn = 'IN (SELECT sheet_package_id FROM #rqcf_pkg)'
        $scopeDocIn = 'IN (SELECT document_guid FROM #rqcf_doc)'
        $commentRunSub = @"
SELECT r.run_id FROM qc_comment_runs r
WHERE r.document_id $scopeDocIn
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @folderLike
"@

        $dbCounts = [ordered]@{}
        if (-not $SkipPreviewCounts.IsPresent) {
            Write-Host '[Database] Counting folder-scoped telemetry rows...' -ForegroundColor Cyan
            $dbCounts.notification_log = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM notification_log n
WHERE n.sheet_package_id $scopePkgIn OR ($nlFolderClause) OR n.document_guid $scopeDocIn
"@
            $dbCounts.document_state_history = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM document_state_history h
WHERE h.sheet_package_id $scopePkgIn OR ($histFolderClause) OR h.document_guid $scopeDocIn
"@
            $dbCounts.transition_events = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM transition_events t
WHERE t.sheet_package_id $scopePkgIn OR ($trFolderClause) OR t.document_guid $scopeDocIn
"@
            $dbCounts.document_activity = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM document_activity d
WHERE d.document_guid $scopeDocIn OR ($actFolderClause)
"@
            $dbCounts.qc_workflow_events = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_workflow_events w
WHERE w.sheet_package_id $scopePkgIn OR w.document_id $scopeDocIn
"@
            $dbCounts.qc_cycle_completions = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_cycle_completions c
WHERE c.sheet_package_id $scopePkgIn OR CAST(c.document_guid AS NVARCHAR(50)) $scopeDocIn
"@
            $dbCounts.audit_events = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM audit_events a
WHERE ($auditFolderClause) OR a.pw_objguid $scopeDocIn
"@
            $dbCounts.processing_jobs = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM processing_jobs j
WHERE ($jobFolderClause) OR j.sheet_package_id $scopePkgIn
"@
            if ($IncludeCommentTelemetry.IsPresent) {
                $dbCounts.qc_comment_status_history = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_comment_status_history h
WHERE h.document_id $scopeDocIn OR h.detected_run_id IN ($commentRunSub)
"@
                $dbCounts.qc_comments = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_comments c WHERE c.run_id IN ($commentRunSub)
"@
                $dbCounts.qc_comment_runs = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM qc_comment_runs r
WHERE r.document_id $scopeDocIn
   OR REPLACE(LOWER(LTRIM(RTRIM(ISNULL(r.pw_path, '')))), '\', '/') LIKE @folderLike
"@
            }
        } else {
            Write-Host '[Database] Preview counts skipped (-SkipPreviewCounts).' -ForegroundColor DarkGray
        }

        $sheetPackagesMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM #rqcf_pkg
"@
        $sheetDocumentsMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_documents sd WHERE sd.sheet_package_id $scopePkgIn
"@

        $sheetIndexMatched = 0L
        $sheetIndexCompletionData = 0L
        $sheetPackageCompletionData = 0L
        if (-not $SkipSheetIndexUpdate.IsPresent) {
            $sheetIndexMatched = _RQCF-GetScalarConn -Connection $dbConn -Params $params -CommandTimeout $QueryTimeoutSeconds -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si WHERE ($siFolderClause)
"@
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
            sheetIndexCompletionData = $sheetIndexCompletionData
            sheetPackagesMatched = $sheetPackagesMatched
            sheetDocumentsMatched = $sheetDocumentsMatched
            sheetPackageCompletionData = $sheetPackageCompletionData
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
                $deleted.qc_comment_status_history = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_comment_status_history' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_comment_status_history
WHERE history_id IN (
  SELECT h.history_id FROM qc_comment_status_history h
  WHERE h.document_id $scopeDocIn OR h.detected_run_id IN ($commentRunSub)
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

            $deleted.notification_log = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'notification_log' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM notification_log
WHERE id IN (
  SELECT n.id FROM notification_log n
  WHERE n.sheet_package_id $scopePkgIn OR ($nlFolderClause) OR n.document_guid $scopeDocIn
)
"@
            $deleted.qc_workflow_events = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_workflow_events' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_workflow_events
WHERE event_id IN (
  SELECT w.event_id FROM qc_workflow_events w
  WHERE w.sheet_package_id $scopePkgIn OR w.document_id $scopeDocIn
)
"@
            $deleted.qc_cycle_completions = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'qc_cycle_completions' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM qc_cycle_completions
WHERE id IN (
  SELECT c.id FROM qc_cycle_completions c
  WHERE c.sheet_package_id $scopePkgIn OR CAST(c.document_guid AS NVARCHAR(50)) $scopeDocIn
)
"@
            $deleted.transition_events = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'transition_events' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM transition_events
WHERE id IN (
  SELECT t.id FROM transition_events t
  WHERE t.sheet_package_id $scopePkgIn OR ($trFolderClause) OR t.document_guid $scopeDocIn
)
"@
            $deleted.document_state_history = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'document_state_history' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM document_state_history
WHERE id IN (
  SELECT h.id FROM document_state_history h
  WHERE h.sheet_package_id $scopePkgIn OR ($histFolderClause) OR h.document_guid $scopeDocIn
)
"@
            $deleted.document_activity = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'document_activity' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM document_activity
WHERE id IN (
  SELECT d.id FROM document_activity d
  WHERE d.document_guid $scopeDocIn OR ($actFolderClause)
)
"@
            $deleted.audit_events = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'audit_events' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM audit_events
WHERE id IN (
  SELECT a.id FROM audit_events a
  WHERE ($auditFolderClause) OR a.pw_objguid $scopeDocIn
)
"@
            $deleted.processing_jobs = _RQCF-RunDeleteLoopConn -Connection $dbConn -Params $params -Label 'processing_jobs' -CommandTimeout $QueryTimeoutSeconds -Sql @"
DELETE TOP (@batchSize) FROM processing_jobs
WHERE id IN (
  SELECT j.id FROM processing_jobs j
  WHERE ($jobFolderClause) OR j.sheet_package_id $scopePkgIn
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
                    $summary.sheetIndexUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $sheetSql -Params $params -CommandTimeout $QueryTimeoutSeconds
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
                    $summary.sheetPackagesUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $packageSql -Params $params -CommandTimeout $QueryTimeoutSeconds
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
                    $summary.sheetDocumentsUpdated = _RQCF-RunNonQueryConn -Connection $dbConn -Sql $docSql -Params $params -CommandTimeout $QueryTimeoutSeconds
                    Write-Host ('  [sheet_documents] updated {0} row(s): pw_state_name reset (not deleted).' -f $summary.sheetDocumentsUpdated) -ForegroundColor Green
                }
            } else {
                Write-Host '  [sheet_index / sheet_packages / sheet_documents] skipped (-SkipSheetIndexUpdate).' -ForegroundColor DarkGray
            }

            Write-Host ("Done. Deleted {0} telemetry row(s)." -f $summary.totalRowsDeleted) -ForegroundColor Green
        } else {
            Write-Host 'Dry run: no database rows deleted. Pass -ConfirmReset to apply.' -ForegroundColor Yellow
        }
    } finally {
        try { $sessionRes.Data.session.Dispose() } catch { }
    }
} else {
    Write-Host '[Database] Skipped (-SkipDatabase).' -ForegroundColor DarkGray
    $summary.database = @{ skipped = $true }
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }
