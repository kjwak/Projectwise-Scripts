# Lane QC PDF DOCUMENT_DELETE: registry purge + qc_workflow_events telemetry.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Core.Database.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-True($v, $msg) { if (-not $v) { throw "ASSERT FAILED: $msg" } }

$config = @{ database = @{ enabled = $true; connectionString = 'x' } }

InModuleScope -ModuleName Core.Database {
    $script:nonQuerySql = @()
    $script:workflowEvent = $null

    function Test-QCDatabaseWritesAllowed { param([hashtable]$Config) return $true }
    function _QDB-IsEnabled { param([hashtable]$Config) return $true }
    function _QDB-NormalizeTelemetryPath { param([string]$Path) return $Path }
    function Get-SheetPackageIdForDocument { param([hashtable]$Config, [string]$DocumentGuid) return [guid]'11111111-1111-1111-1111-111111111111' }

    function Invoke-QCDatabaseQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        if ($Sql -match 'last_state') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('last_state', [string])
            [void]$table.Rows.Add('Redlines Received')
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
    }

    function Invoke-QCDatabaseNonQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        $script:nonQuerySql += $Sql
        return New-QCSuccessResult -Code 'DB_NONQUERY_OK' -Message 'ok' -Data @{ rowsAffected = 1 }
    }

    function Write-QCWorkflowEventRow {
        param(
            [hashtable]$Config,
            [string]$DocumentId,
            [string]$JobId,
            [string]$EventType,
            [string]$PreviousPwState,
            [string]$TargetPwState,
            [string]$DecisionCode,
            [string]$QcReviewType,
            [string]$PayloadJson,
            [guid]$SheetPackageId,
            [switch]$PlannedOnly
        )
        $script:workflowEvent = [pscustomobject]@{
            DocumentId      = $DocumentId
            JobId           = $JobId
            EventType       = $EventType
            PreviousPwState = $PreviousPwState
            DecisionCode    = $DecisionCode
            QcReviewType    = $QcReviewType
            PayloadJson     = $PayloadJson
            SheetPackageId  = $SheetPackageId
            PlannedOnly     = [bool]$PlannedOnly
        }
        return New-QCSuccessResult -Code 'QC_WORKFLOW_EVENT_WRITTEN' -Message 'ok' -Data @{ written = $true }
    }

    $script:nonQuerySql = @()
    $script:workflowEvent = $null

    $res = Remove-QCLaneQcPdfRegistryRecords -Config $config -DocumentGuid 'da521122-0000-4000-8000-000000000001' `
        -DocumentName '080J082001ca001-chk.pdf' -FolderPath 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1' -AuditEventId 42103

    Assert-True $res.IsSuccess 'delete cleanup should succeed'
    Assert-Eq $res.Code 'LANE_PDF_DELETE_PROCESSED' 'success code'
    Assert-Eq $script:nonQuerySql.Count 5 'five registry SQL statements'

    $joined = ($script:nonQuerySql -join ' ')
    Assert-True ($joined -match 'sheet_package_qc_pdfs') 'deactivates sheet_package_qc_pdfs'
    Assert-True ($joined -match 'DELETE FROM sheet_index') 'removes sheet_index row'
    Assert-True ($joined -match 'qc_pdf_guid = NULL') 'clears stem qc_pdf links'
    Assert-True ($joined -match 'sheet_packages') 'clears sheet_packages lane columns'
    Assert-True ($joined -match 'DELETE FROM sheet_documents') 'removes sheet_documents row'

    Assert-True ($null -ne $script:workflowEvent) 'workflow event should be written'
    Assert-Eq $script:workflowEvent.EventType 'DOCUMENT_DELETE' 'event type mirrors audit action'
    Assert-Eq $script:workflowEvent.DecisionCode 'QC_LANE_PDF_REGISTRY_PURGED' 'decision code'
    Assert-Eq $script:workflowEvent.QcReviewType 'check' 'lane process type'
    Assert-Eq $script:workflowEvent.PreviousPwState 'Redlines Received' 'captures last known state'
    Assert-Eq $script:workflowEvent.JobId 'qc_lane_delete_42103' 'job id from audit event'

    $skip = Remove-QCLaneQcPdfRegistryRecords -Config $config -DocumentGuid 'da521122-0000-4000-8000-000000000001' `
        -DocumentName '080J082001ca001.pdf' -FolderPath 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
    Assert-Eq $skip.Code 'LANE_PDF_DELETE_SKIPPED' 'stem PDF delete is skipped'
}

# Watch script references lane delete handler.
$watchText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts/Watch-QCTrigger.ps1') -Raw
Assert-True ($watchText -match 'Remove-QCLaneQcPdfRegistryRecords') 'Watch-QCTrigger calls Remove-QCLaneQcPdfRegistryRecords'
Assert-True ($watchText -match 'DOCUMENT_DELETE') 'Watch-QCTrigger handles DOCUMENT_DELETE'

Write-Host 'OK: lane PDF delete registry tests passed.' -ForegroundColor Green
