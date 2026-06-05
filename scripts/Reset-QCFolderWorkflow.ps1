<#
.SYNOPSIS
Resets ProjectWise workflow state for all documents in a folder and clears QC telemetry in SQL.

.DESCRIPTION
For a single Sheets (or any) folder:

1. ProjectWise — sets workflow state on every document in the folder to the target state
   (default: qcWorkflow.states.production, usually "In Production").
2. Database — deletes folder-scoped rows from QC telemetry tables and refreshes sheet_index.

Default is preview only. Pass -ConfirmReset to apply PW and database changes.

Does not delete sheet_index rows or ProjectWise documents. Does not remove queue JSON jobs
(use Purge-QCPendingByFilters or manual queue cleanup separately).

.EXAMPLE
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath 'AZFWY1704-FD02-SR202\CADD\Sheets' -ConfirmReset
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipProjectWise
.\scripts\Reset-QCFolderWorkflow.ps1 -FolderPath '...\Sheets' -ConfirmReset -SkipDatabase
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
$folderPathClause = _RQCF-PathLikeClause -ColumnSql 'folder_path' -Params $params -ParamName 'folderLike'
$resolvedFolderClause = _RQCF-PathLikeClause -ColumnSql 'resolved_folder' -Params $params -ParamName 'folderLike'
$sourceFolderClause = _RQCF-PathLikeClause -ColumnSql 'source_folder' -Params $params -ParamName 'folderLike'
$guidSub = _RQCF-SheetIndexGuidSubquery -FolderLikeClause $siFolderClause

$summary = [ordered]@{
    dryRun            = $DryRun.IsPresent
    folderPath        = $normFolder
    targetState       = $TargetState
    projectWise       = $null
    database          = $null
    totalRowsDeleted  = 0
    sheetIndexUpdated = 0
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

    Write-Host '[ProjectWise] Scanning folder documents...' -ForegroundColor Cyan
    $pwResult = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -ScriptBlock {
        $docs = @(Get-PWDocumentsInFolderRaw -FolderPath $using:apiPath)
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
            $needs = -not ($cur -ieq $using:TargetState)
            if ($needs) { $needsChange++ } else { $alreadyTarget++ }
            if ($samples.Count -lt 8) {
                $samples.Add([pscustomobject]@{ name = $name; currentState = $cur; needsChange = $needs }) | Out-Null
            }
            if (-not $using:doApply -or -not $needs) { continue }
            try {
                _RQCF-SetPwDocumentState -Document $doc -StateName $using:TargetState
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
WHERE document_guid IN ($guidSub) OR ($folderPathClause)
"@
    $dbCounts.document_state_history = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM document_state_history
WHERE document_guid IN ($guidSub) OR ($folderPathClause)
"@
    $dbCounts.transition_events = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM transition_events
WHERE document_guid IN ($guidSub) OR ($folderPathClause)
"@
    $dbCounts.document_activity = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM document_activity
WHERE document_guid IN ($guidSub) OR ($folderPathClause)
"@
    $dbCounts.qc_workflow_events = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM qc_workflow_events w
WHERE w.document_id IN ($guidSub)
"@
    $dbCounts.audit_events = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM audit_events
WHERE ($resolvedFolderClause) OR pw_objguid IN ($guidSub)
"@
    $dbCounts.processing_jobs = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM processing_jobs
WHERE ($sourceFolderClause)
"@
    $dbCounts.sheet_index_rows = _RQCF-GetScalar -Config $config -Params $params -Sql @"
SELECT COUNT_BIG(1) FROM sheet_index si WHERE ($siFolderClause)
"@

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

    $summary.database = @{ rowsMatched = $dbCounts; deleted = [ordered]@{} }
    foreach ($k in @($dbCounts.Keys)) {
        Write-Host ("  {0}: {1}" -f $k, $dbCounts[$k]) -ForegroundColor Gray
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
  WHERE n.document_guid IN ($guidSub) OR ($folderPathClause)
)
"@
        $deleted.qc_workflow_events = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'qc_workflow_events' -Sql @"
DELETE TOP (@batchSize) FROM qc_workflow_events
WHERE event_id IN (SELECT w.event_id FROM qc_workflow_events w WHERE w.document_id IN ($guidSub))
"@
        $deleted.transition_events = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'transition_events' -Sql @"
DELETE TOP (@batchSize) FROM transition_events
WHERE id IN (
  SELECT t.id FROM transition_events t
  WHERE t.document_guid IN ($guidSub) OR ($folderPathClause)
)
"@
        $deleted.document_state_history = _RQCF-RunDeleteLoop -Config $config -Params $params -Label 'document_state_history' -Sql @"
DELETE TOP (@batchSize) FROM document_state_history
WHERE id IN (
  SELECT h.id FROM document_state_history h
  WHERE h.document_guid IN ($guidSub) OR ($folderPathClause)
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
WHERE id IN (SELECT j.id FROM processing_jobs j WHERE ($sourceFolderClause))
"@

        foreach ($k in @($deleted.Keys)) {
            $summary.totalRowsDeleted += [long]$deleted[$k]
        }

        $sheetSql = @"
UPDATE sheet_index
SET pw_state_name = @targetState,
    last_updated_at = SYSDATETIMEOFFSET()
"@
        if (-not $KeepSheetIndexQcFields.IsPresent) {
            $sheetSql += "`r`n    qc_stage = NULL,`r`n    qc_status = NULL,`r`n    last_audit_event_at = NULL"
        }
        $sheetSql += "`r`nWHERE ($siFolderClause)"

        if ($PSCmdlet.ShouldProcess('sheet_index', 'UPDATE pw_state_name')) {
            $upd = Invoke-QCDatabaseNonQuery -Config $config -Sql $sheetSql -Parameters $params
            if (-not $upd.IsSuccess) { throw $upd.Message }
            $summary.sheetIndexUpdated = [int]$upd.Data.rowsAffected
            Write-Host ('  [sheet_index] updated {0} row(s).' -f $summary.sheetIndexUpdated) -ForegroundColor Green
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
