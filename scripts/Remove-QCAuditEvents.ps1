<#
.SYNOPSIS
Removes aged rows from dbo.audit_events (ingested ProjectWise audit trail copies).

.DESCRIPTION
audit_events can grow quickly. Trim by capture time (when the poller stored the row).
By default only deletes rows with processed = 1 so unprocessed backlog is kept.

Default is preview; pass -ConfirmDeletes to apply. Uses batched DELETE TOP for large tables.

.EXAMPLE
.\scripts\Remove-QCAuditEvents.ps1 -OlderThanDays 90
.\scripts\Remove-QCAuditEvents.ps1 -ConfirmDeletes -OlderThanDays 60 -ProcessedOnly
.\scripts\Remove-QCAuditEvents.ps1 -ConfirmDeletes -OlderThanDays 30 -IncludeUnprocessed
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [int]$OlderThanDays = 90,
    [datetime]$BeforeUtc,
    [switch]$ProcessedOnly,
    [switch]$IncludeUnprocessed,
    [string]$ResolvedFolderLike = '',
    [int]$BatchSize = 5000,
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
if ($ProcessedOnly.IsPresent -and $IncludeUnprocessed.IsPresent) {
    throw 'Use -ProcessedOnly or -IncludeUnprocessed, not both.'
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

$onlyProcessed = $true
if ($IncludeUnprocessed.IsPresent) { $onlyProcessed = $false }
elseif ($ProcessedOnly.IsPresent) { $onlyProcessed = $true }

$cutoff = $null
if ($TruncateAll.IsPresent) {
    $cutoff = $null
} elseif ($BeforeUtc) {
    $cutoff = [datetimeoffset]$BeforeUtc.ToUniversalTime()
} else {
    $cutoff = [datetimeoffset]::UtcNow.AddDays(-1 * $OlderThanDays)
}

function _QAE-BuildWhereClause {
    param(
        [bool]$UseCutoff,
        [hashtable]$Params
    )
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($UseCutoff) {
        $Params['cutoff'] = $cutoff
        $parts.Add('captured_at < @cutoff') | Out-Null
    }
    if ($onlyProcessed) {
        $parts.Add('processed = 1') | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($ResolvedFolderLike)) {
        $Params['folderLike'] = $ResolvedFolderLike.Trim()
        $parts.Add('resolved_folder LIKE @folderLike') | Out-Null
    }
    if ($parts.Count -eq 0) { return '' }
    return ' WHERE ' + ($parts -join ' AND ')
}

$params = @{ batchSize = $BatchSize }
$useCutoff = -not $TruncateAll.IsPresent
$where = _QAE-BuildWhereClause -UseCutoff $useCutoff -Params $params
$countSql = "SELECT COUNT_BIG(1) AS cnt FROM audit_events$where"
$deleteSql = "DELETE TOP (@batchSize) FROM audit_events$where"

Write-Host '[audit_events] Counting matching rows...' -ForegroundColor Cyan
$countRes = Invoke-QCDatabaseScalar -Config $config -Sql $countSql -Parameters $params
if (-not $countRes.IsSuccess) { throw $countRes.Message }
$matched = [long]$countRes.Data.value

$summary = [ordered]@{
    dryRun              = $DryRun.IsPresent
    truncateAll         = $TruncateAll.IsPresent
    cutoffUtc           = if ($cutoff) { $cutoff.UtcDateTime.ToString('o') } else { $null }
    olderThanDays       = if ($OlderThanDays -gt 0) { $OlderThanDays } else { $null }
    processedOnly       = $onlyProcessed
    resolvedFolderLike  = $ResolvedFolderLike
    rowsMatched         = $matched
    deleted             = 0
}

Write-Host ("  Matched {0} row(s) (processedOnly={1})." -f $matched, $onlyProcessed) -ForegroundColor $(if ($matched -eq 0) { 'Green' } else { 'Yellow' })

if ($doDelete -and $matched -gt 0) {
    $total = 0L
    do {
        $batch = 0
        if ($PSCmdlet.ShouldProcess('audit_events', 'DELETE batch')) {
            $delRes = Invoke-QCDatabaseNonQuery -Config $config -Sql $deleteSql -Parameters $params
            if (-not $delRes.IsSuccess) { throw $delRes.Message }
            $batch = [int]$delRes.Data.rowsAffected
            $total += $batch
            if ($batch -gt 0) {
                Write-Host ("  Deleted {0} row(s), running total {1}" -f $batch, $total) -ForegroundColor Gray
            }
        } else {
            break
        }
    } while ($batch -ge $BatchSize)
    $summary.deleted = $total
    Write-Host ("Done. Deleted {0} audit_events row(s)." -f $total) -ForegroundColor Green
} elseif (-not $doDelete) {
    Write-Host 'Dry run: no rows deleted. Pass -ConfirmDeletes to apply.' -ForegroundColor Yellow
} else {
    Write-Host 'Nothing to delete.' -ForegroundColor Green
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 5 }
