<#
.SYNOPSIS
Removes rows from dbo.qc_workflow_events (and optional comment-sync tables).

.DESCRIPTION
-TruncateAll deletes every row in qc_workflow_events (no folder_path on that table).
Optional -IncludeCommentTelemetry also clears qc_comment_runs, qc_comments, and
qc_comment_status_history.

For scoped cleanup, -FolderPathLike resolves document_id via sheet_index.folder_path.
Default is preview; pass -ConfirmDeletes to apply.

.EXAMPLE
.\scripts\Remove-QCWorkflowEvents.ps1 -TruncateAll
.\scripts\Remove-QCWorkflowEvents.ps1 -TruncateAll -IncludeCommentTelemetry -ConfirmDeletes
.\scripts\Remove-QCWorkflowEvents.ps1 -ConfirmDeletes -FolderPathLike 'AZFWY2302-018'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [string[]]$FolderPathLike = @(),
    [string]$FolderPathFilter = '',
    [string]$DocumentId = '',
    [string]$JobId = '',
    [int]$OlderThanDays = 0,
    [datetime]$BeforeUtc,
    [int]$BatchSize = 5000,
    [switch]$IncludeCommentTelemetry,
    [Alias('All')]
    [switch]$TruncateAll,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Import-QCScriptModules.ps1') -RepoRoot $repoRoot

if ($DryRun.IsPresent -and $ConfirmDeletes.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmDeletes (apply deletes), not both.'
}
$doDelete = $ConfirmDeletes.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmDeletes.IsPresent) {
    Write-Host 'Preview only: pass -ConfirmDeletes to delete rows, or -DryRun explicitly.' -ForegroundColor Yellow
    $DryRun = $true
}

$pathPatterns = [System.Collections.Generic.List[string]]::new()
foreach ($p in @($FolderPathLike)) {
    if (-not [string]::IsNullOrWhiteSpace($p)) { $pathPatterns.Add($p.Trim()) | Out-Null }
}
if (-not [string]::IsNullOrWhiteSpace($FolderPathFilter)) {
    $pathPatterns.Add($FolderPathFilter.Trim()) | Out-Null
}

$hasFolderScope = $pathPatterns.Count -gt 0
$hasAgeScope = $TruncateAll.IsPresent -or $OlderThanDays -gt 0 -or $BeforeUtc
$hasIdScope = -not [string]::IsNullOrWhiteSpace($DocumentId) -or -not [string]::IsNullOrWhiteSpace($JobId)

if (-not $TruncateAll.IsPresent -and -not $hasFolderScope -and -not $hasIdScope -and -not $hasAgeScope) {
    throw 'Specify -FolderPathLike (path fragment(s)), -DocumentId, -JobId, optional age filters, or -TruncateAll.'
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true.'
}
Initialize-QCDatabaseSchema -Config $config | Out-Null

$cutoff = $null
$useCutoff = $false
if ($TruncateAll.IsPresent) {
    $useCutoff = $false
} elseif ($BeforeUtc) {
    $cutoff = [datetimeoffset]$BeforeUtc.ToUniversalTime()
    $useCutoff = $true
} elseif ($OlderThanDays -gt 0) {
    $cutoff = [datetimeoffset]::UtcNow.AddDays(-1 * $OlderThanDays)
    $useCutoff = $true
}

$fullWipe = $TruncateAll.IsPresent -and -not $hasFolderScope -and -not $hasIdScope -and -not $useCutoff

function _QWE-NormalizePathPattern {
    param([string]$Pattern)
    $norm = ($Pattern -replace '\\', '/').ToLowerInvariant()
    if ($norm -notmatch '[%_\[]') { return "%$norm%" }
    return $norm
}

function _QWE-PathLikeClauses {
    param(
        [string]$ColumnSql,
        [System.Collections.Generic.List[string]]$Patterns,
        [hashtable]$Params,
        [string]$Prefix
    )
    $clauses = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $Patterns.Count; $i++) {
        $key = '{0}{1}' -f $Prefix, $i
        $Params[$key] = _QWE-NormalizePathPattern -Pattern $Patterns[$i]
        $clauses.Add(
            "REPLACE(LOWER(LTRIM(RTRIM(ISNULL($ColumnSql, '')))), '\', '/') LIKE @$key"
        ) | Out-Null
    }
    return $clauses
}

function _QWE-SheetIndexGuidSubquery {
    param([hashtable]$Params, [string]$Prefix)
    $clauses = _QWE-PathLikeClauses -ColumnSql 'si.folder_path' -Patterns $pathPatterns -Params $Params -Prefix $Prefix
    return @"
SELECT LTRIM(RTRIM(si.document_guid))
FROM sheet_index si
WHERE si.document_guid IS NOT NULL AND (($($clauses -join ' OR ')))
"@
}

function _QWE-BuildWorkflowWhere {
    param([hashtable]$Params)
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($useCutoff) {
        $Params['cutoff'] = $cutoff
        $parts.Add('w.created_utc < @cutoff') | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($DocumentId)) {
        $Params['documentId'] = $DocumentId.Trim()
        $parts.Add('w.document_id = @documentId') | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($JobId)) {
        $Params['jobId'] = $JobId.Trim()
        $parts.Add('w.job_id = @jobId') | Out-Null
    }
    if ($hasFolderScope) {
        $parts.Add('w.document_id IN (' + (_QWE-SheetIndexGuidSubquery -Params $Params -Prefix 'wf') + ')') | Out-Null
    }
    if ($parts.Count -eq 0) { return '' }
    return ' WHERE ' + ($parts -join ' AND ')
}

function _QWE-BuildCommentRunWhere {
    param([hashtable]$Params)
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($useCutoff) {
        $Params['cutoff'] = $cutoff
        $parts.Add('r.created_utc < @cutoff') | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($DocumentId)) {
        $Params['documentId'] = $DocumentId.Trim()
        $parts.Add('r.document_id = @documentId') | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($JobId)) {
        $Params['jobId'] = $JobId.Trim()
        $parts.Add('r.job_id = @jobId') | Out-Null
    }
    if ($hasFolderScope) {
        $pathClauses = _QWE-PathLikeClauses -ColumnSql 'r.pw_path' -Patterns $pathPatterns -Params $Params -Prefix 'cr'
        $parts.Add(@"
(
    (($($pathClauses -join ' OR ')))
    OR r.document_id IN ($(_QWE-SheetIndexGuidSubquery -Params $Params -Prefix 'crs'))
)
"@) | Out-Null
    }
    if ($parts.Count -eq 0) { return '' }
    return ' WHERE ' + ($parts -join ' AND ')
}

function _QWE-GetScalar {
    param([hashtable]$Config, [string]$Sql, [hashtable]$Params)
    $res = Invoke-QCDatabaseScalar -Config $Config -Sql $Sql -Parameters $Params
    if (-not $res.IsSuccess) { throw $res.Message }
    return [long]$res.Data.value
}

function _QWE-RunDeleteLoop {
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

$params = @{ batchSize = $BatchSize }
$wfWhere = _QWE-BuildWorkflowWhere -Params $params
$runWhere = _QWE-BuildCommentRunWhere -Params $params
$runSub = "SELECT r.run_id FROM qc_comment_runs r$runWhere"

$histWhereParts = [System.Collections.Generic.List[string]]::new()
$histWhereParts.Add("h.detected_run_id IN ($runSub)") | Out-Null
if ($hasFolderScope) {
    $histWhereParts.Add("h.document_id IN ($(_QWE-SheetIndexGuidSubquery -Params $params -Prefix 'hs'))") | Out-Null
}
$histWhere = ' WHERE ' + ($histWhereParts -join ' OR ')

$summary = [ordered]@{
    dryRun                  = $DryRun.IsPresent
    truncateAll             = $TruncateAll.IsPresent
    folderPathLike          = @($pathPatterns)
    documentId              = $DocumentId
    jobId                   = $JobId
    includeCommentTelemetry = $IncludeCommentTelemetry.IsPresent
    qc_workflow_events      = $null
    commentTelemetry        = $null
    totalDeleted            = 0
}

Write-Host '[qc_workflow_events] Counting matching rows...' -ForegroundColor Cyan
if ($fullWipe) {
    $wfCount = _QWE-GetScalar -Config $config -Sql 'SELECT COUNT_BIG(1) FROM qc_workflow_events' -Params $params
} else {
    $wfCount = _QWE-GetScalar -Config $config -Sql "SELECT COUNT_BIG(1) FROM qc_workflow_events w$wfWhere" -Params $params
}
$summary.qc_workflow_events = @{ rowsMatched = $wfCount; deleted = 0 }
Write-Host ("  Matched {0} row(s)." -f $wfCount) -ForegroundColor $(if ($wfCount -eq 0) { 'Green' } else { 'Yellow' })
if ($pathPatterns.Count -gt 0) {
    Write-Host ("  Folder patterns: {0}" -f ($pathPatterns -join ', ')) -ForegroundColor Gray
}

if ($IncludeCommentTelemetry.IsPresent) {
    Write-Host '[comment telemetry] Counting matching rows...' -ForegroundColor Cyan
    if ($fullWipe) {
        $runCount = _QWE-GetScalar -Config $config -Sql 'SELECT COUNT_BIG(1) FROM qc_comment_runs' -Params $params
        $commentCount = _QWE-GetScalar -Config $config -Sql 'SELECT COUNT_BIG(1) FROM qc_comments' -Params $params
        $histCount = _QWE-GetScalar -Config $config -Sql 'SELECT COUNT_BIG(1) FROM qc_comment_status_history' -Params $params
    } else {
        $runCount = _QWE-GetScalar -Config $config -Sql "SELECT COUNT_BIG(1) FROM qc_comment_runs r$runWhere" -Params $params
        $commentCount = _QWE-GetScalar -Config $config -Sql "SELECT COUNT_BIG(1) FROM qc_comments c WHERE c.run_id IN ($runSub)" -Params $params
        $histCount = _QWE-GetScalar -Config $config -Sql "SELECT COUNT_BIG(1) FROM qc_comment_status_history h$histWhere" -Params $params
    }
    $summary.commentTelemetry = @{
        rowsMatched = @{ runs = $runCount; comments = $commentCount; history = $histCount }
        deleted     = $null
    }
    Write-Host ("  qc_comment_runs: {0}" -f $runCount) -ForegroundColor Gray
    Write-Host ("  qc_comments: {0}" -f $commentCount) -ForegroundColor Gray
    Write-Host ("  qc_comment_status_history: {0}" -f $histCount) -ForegroundColor Gray
}

if ($doDelete) {
    if ($IncludeCommentTelemetry.IsPresent) {
        $summary.commentTelemetry.deleted = [ordered]@{}
        if ($fullWipe) {
            $summary.commentTelemetry.deleted.qc_comments = _QWE-RunDeleteLoop -Config $config `
                -Sql 'DELETE TOP (@batchSize) FROM qc_comments' -Params $params -Label 'qc_comments'
            $summary.commentTelemetry.deleted.qc_comment_status_history = _QWE-RunDeleteLoop -Config $config `
                -Sql 'DELETE TOP (@batchSize) FROM qc_comment_status_history' -Params $params -Label 'qc_comment_status_history'
            $summary.commentTelemetry.deleted.qc_comment_runs = _QWE-RunDeleteLoop -Config $config `
                -Sql 'DELETE TOP (@batchSize) FROM qc_comment_runs' -Params $params -Label 'qc_comment_runs'
        } else {
            $summary.commentTelemetry.deleted.qc_comment_status_history = _QWE-RunDeleteLoop -Config $config `
                -Sql "DELETE TOP (@batchSize) FROM qc_comment_status_history WHERE history_id IN (SELECT h.history_id FROM qc_comment_status_history h$histWhere)" `
                -Params $params -Label 'qc_comment_status_history'
            $summary.commentTelemetry.deleted.qc_comments = _QWE-RunDeleteLoop -Config $config `
                -Sql "DELETE TOP (@batchSize) FROM qc_comments WHERE comment_record_id IN (SELECT c.comment_record_id FROM qc_comments c WHERE c.run_id IN ($runSub))" `
                -Params $params -Label 'qc_comments'
            $summary.commentTelemetry.deleted.qc_comment_runs = _QWE-RunDeleteLoop -Config $config `
                -Sql "DELETE TOP (@batchSize) FROM qc_comment_runs WHERE run_id IN (SELECT r.run_id FROM qc_comment_runs r$runWhere)" `
                -Params $params -Label 'qc_comment_runs'
        }
        $summary.totalDeleted += [long]$summary.commentTelemetry.deleted.qc_comment_status_history
        $summary.totalDeleted += [long]$summary.commentTelemetry.deleted.qc_comments
        $summary.totalDeleted += [long]$summary.commentTelemetry.deleted.qc_comment_runs
    }

    if ($fullWipe) {
        $summary.qc_workflow_events.deleted = _QWE-RunDeleteLoop -Config $config `
            -Sql 'DELETE TOP (@batchSize) FROM qc_workflow_events' -Params $params -Label 'qc_workflow_events'
    } else {
        $summary.qc_workflow_events.deleted = _QWE-RunDeleteLoop -Config $config `
            -Sql "DELETE TOP (@batchSize) FROM qc_workflow_events WHERE event_id IN (SELECT w.event_id FROM qc_workflow_events w$wfWhere)" `
            -Params $params -Label 'qc_workflow_events'
    }
    $summary.totalDeleted += [long]$summary.qc_workflow_events.deleted
    Write-Host ("Done. Total deleted: {0}" -f $summary.totalDeleted) -ForegroundColor Green
} else {
    Write-Host 'Dry run: no rows deleted. Pass -ConfirmDeletes to apply.' -ForegroundColor Yellow
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }
