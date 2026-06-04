<#
.SYNOPSIS
Removes rows from dbo.qc_workflow_events (workflow transition telemetry).

.DESCRIPTION
qc_workflow_events is debug/replay telemetry, not execution state. Safe to trim by age,
document, or job. Default is preview; pass -ConfirmDeletes to apply.

Optional -IncludeCommentTelemetry also removes qc_comment_runs and dependent rows
(qc_comments, qc_comment_status_history) using the same age/filter rules on created_utc.

.EXAMPLE
.\scripts\Remove-QCWorkflowEvents.ps1 -OlderThanDays 180
.\scripts\Remove-QCWorkflowEvents.ps1 -ConfirmDeletes -OlderThanDays 90 -DocumentId '{guid}'
.\scripts\Remove-QCWorkflowEvents.ps1 -ConfirmDeletes -TruncateAll
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [int]$OlderThanDays = 0,
    [string]$DocumentId = '',
    [string]$JobId = '',
    [datetime]$BeforeUtc,
    [int]$BatchSize = 5000,
    [switch]$IncludeCommentTelemetry,
    [switch]$TruncateAll,
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

if ($DryRun.IsPresent -and $ConfirmDeletes.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmDeletes (apply deletes), not both.'
}
$doDelete = $ConfirmDeletes.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmDeletes.IsPresent) {
    Write-Host 'Preview only: pass -ConfirmDeletes to delete rows, or -DryRun explicitly.' -ForegroundColor Yellow
    $DryRun = $true
}

if (-not $TruncateAll.IsPresent) {
    if ($OlderThanDays -le 0 -and -not $BeforeUtc) {
        throw 'Specify -OlderThanDays (>0), -BeforeUtc, or -TruncateAll.'
    }
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true.'
}
Initialize-QCDatabaseSchema -Config $config | Out-Null

$cutoff = $null
if ($TruncateAll.IsPresent) {
    $cutoff = $null
} elseif ($BeforeUtc) {
    $cutoff = [datetimeoffset]$BeforeUtc.ToUniversalTime()
} else {
    $cutoff = [datetimeoffset]::UtcNow.AddDays(-1 * $OlderThanDays)
}

function _QWE-BuildWhereClause {
    param(
        [string]$DateColumn,
        [bool]$UseCutoff,
        [hashtable]$Params
    )
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($UseCutoff) {
        $Params['cutoff'] = $cutoff
        $parts.Add("$DateColumn < @cutoff") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($DocumentId)) {
        $Params['documentId'] = $DocumentId.Trim()
        $parts.Add('document_id = @documentId') | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($JobId)) {
        $Params['jobId'] = $JobId.Trim()
        $parts.Add('job_id = @jobId') | Out-Null
    }
    if ($parts.Count -eq 0) { return '' }
    return ' WHERE ' + ($parts -join ' AND ')
}

function _QWE-GetCount {
    param(
        [hashtable]$Config,
        [string]$Table,
        [string]$DateColumn,
        [bool]$UseCutoff
    )
    $params = @{ batchSize = $BatchSize }
    $where = _QWE-BuildWhereClause -DateColumn $DateColumn -UseCutoff $UseCutoff -Params $params
    $sql = "SELECT COUNT_BIG(1) AS cnt FROM $Table$where"
    $res = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
    if (-not $res.IsSuccess) { throw $res.Message }
    return [long]$res.Data.value
}

function _QWE-DeleteBatched {
    param(
        [hashtable]$Config,
        [string]$Table,
        [string]$DateColumn,
        [bool]$UseCutoff,
        [string]$Label
    )
    $params = @{ batchSize = $BatchSize }
    $where = _QWE-BuildWhereClause -DateColumn $DateColumn -UseCutoff $UseCutoff -Params $params
    $deleteSql = "DELETE TOP (@batchSize) FROM $Table$where"
    $total = 0L
    do {
        $batch = 0
        if ($PSCmdlet.ShouldProcess($Label, 'DELETE batch')) {
            $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $deleteSql -Parameters $params
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

$summary = [ordered]@{
    dryRun                   = $DryRun.IsPresent
    truncateAll              = $TruncateAll.IsPresent
    cutoffUtc                = if ($cutoff) { $cutoff.UtcDateTime.ToString('o') } else { $null }
    olderThanDays            = if ($OlderThanDays -gt 0) { $OlderThanDays } else { $null }
    documentId               = $DocumentId
    jobId                    = $JobId
    includeCommentTelemetry  = $IncludeCommentTelemetry.IsPresent
    qc_workflow_events       = $null
    qc_comment_status_history = $null
    qc_comments              = $null
    qc_comment_runs          = $null
    totalDeleted             = 0
}

$useCutoff = -not $TruncateAll.IsPresent

Write-Host '[qc_workflow_events] Counting matching rows...' -ForegroundColor Cyan
$wfCount = _QWE-GetCount -Config $config -Table 'qc_workflow_events' -DateColumn 'created_utc' -UseCutoff $useCutoff
$summary.qc_workflow_events = @{ rowsMatched = $wfCount; deleted = 0 }
Write-Host ("  Matched {0} row(s)." -f $wfCount) -ForegroundColor $(if ($wfCount -eq 0) { 'Green' } else { 'Yellow' })

if ($IncludeCommentTelemetry.IsPresent) {
    Write-Host '[comment telemetry] Counting matching rows...' -ForegroundColor Cyan
    $histCount = _QWE-GetCount -Config $config -Table 'qc_comment_status_history' -DateColumn 'detected_utc' -UseCutoff $useCutoff
    $commentCount = _QWE-GetCount -Config $config -Table 'qc_comments' -DateColumn 'inserted_utc' -UseCutoff $useCutoff
    $runCount = _QWE-GetCount -Config $config -Table 'qc_comment_runs' -DateColumn 'created_utc' -UseCutoff $useCutoff
    $summary.qc_comment_status_history = @{ rowsMatched = $histCount; deleted = 0 }
    $summary.qc_comments = @{ rowsMatched = $commentCount; deleted = 0 }
    $summary.qc_comment_runs = @{ rowsMatched = $runCount; deleted = 0 }
    Write-Host ("  qc_comment_status_history: {0}" -f $histCount) -ForegroundColor Gray
    Write-Host ("  qc_comments: {0}" -f $commentCount) -ForegroundColor Gray
    Write-Host ("  qc_comment_runs: {0}" -f $runCount) -ForegroundColor Gray
}

if ($doDelete) {
    if ($IncludeCommentTelemetry.IsPresent) {
        $summary.qc_comment_status_history.deleted = _QWE-DeleteBatched -Config $config -Table 'qc_comment_status_history' -DateColumn 'detected_utc' -UseCutoff $useCutoff -Label 'qc_comment_status_history'
        $summary.qc_comments.deleted = _QWE-DeleteBatched -Config $config -Table 'qc_comments' -DateColumn 'inserted_utc' -UseCutoff $useCutoff -Label 'qc_comments'
        $summary.qc_comment_runs.deleted = _QWE-DeleteBatched -Config $config -Table 'qc_comment_runs' -DateColumn 'created_utc' -UseCutoff $useCutoff -Label 'qc_comment_runs'
        $summary.totalDeleted += [long]$summary.qc_comment_status_history.deleted
        $summary.totalDeleted += [long]$summary.qc_comments.deleted
        $summary.totalDeleted += [long]$summary.qc_comment_runs.deleted
    }
    $summary.qc_workflow_events.deleted = _QWE-DeleteBatched -Config $config -Table 'qc_workflow_events' -DateColumn 'created_utc' -UseCutoff $useCutoff -Label 'qc_workflow_events'
    $summary.totalDeleted += [long]$summary.qc_workflow_events.deleted
    Write-Host ("Done. Total deleted: {0}" -f $summary.totalDeleted) -ForegroundColor Green
} else {
    Write-Host 'Dry run: no rows deleted. Pass -ConfirmDeletes to apply.' -ForegroundColor Yellow
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }
