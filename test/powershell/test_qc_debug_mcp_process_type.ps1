# QC.DebugMcp qc_process_type diagnostics (filename lane suffix vs DB/PW).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Diagnostics\QC.DebugMcp.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-True($c, $msg) { if (-not $c) { throw "ASSERT FAILED: $msg" } }

InModuleScope -ModuleName QC.DebugMcp {
    function _QDM-TestTableExists { param($TableName) return $false }
    function _QDM-LoadSheetPackageQcPdfRows {
        param([string[]]$PackageIds)
        return @(
            @{
                sheet_package_id = 'pkg-1'
                qc_process_type = 'review'
                document_guid = '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4'
                document_name = '080J082001ab001-rev.pdf'
                is_active = $true
            }
        )
    }

    $lookup = @{ folder_path = 'documents\proj\cadd\sheets\seg_1' }
    $members = @{
        sheet_index_rows = @(
            @{
                document_guid = 'ba1a9c32-2adf-4f39-9d0b-9dac20d0a286'
                document_name = '080J082001ab001.pdf'
                qc_process_type = 'Review'
            }
            @{
                document_guid = '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4'
                document_name = '080J082001ab001-rev.pdf'
                qc_process_type = 'Review'
            }
        )
        sheet_documents_rows = @(
            @{
                document_guid = 'ba1a9c32-2adf-4f39-9d0b-9dac20d0a286'
                document_name = '080J082001ab001.pdf'
                document_role = 'sheet_pdf'
            }
            @{
                document_guid = '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4'
                document_name = '080J082001ab001-rev.pdf'
                document_role = 'qc_pdf'
            }
        )
        sheet_packages_rows = @(@{ sheet_package_id = 'pkg-1'; folder_path = 'documents\proj\cadd\sheets\seg_1' })
    }
    $pwMap = @{
        'ba1a9c32-2adf-4f39-9d0b-9dac20d0a286' = 'Review'
        '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4' = 'Review'
    }

    $diag = _QDM-BuildQcProcessTypeDiagnostics -Lookup $lookup -Members $members -PwProcessTypeByGuid $pwMap -ProjectWiseAvailable $true
    Assert-Eq $diag.checks_failed 0 'consistent review lane passes all checks'
    $rev = @($diag.documents | Where-Object { $_.document_name -eq '080J082001ab001-rev.pdf' })[0]
    Assert-Eq $rev.filename_inferred_process_type 'review' '*-rev.pdf infers review'
    Assert-Eq $rev.database_lane_registry_process_type 'review' 'lane registry review'
    Assert-Eq $rev.projectwise_process_type_normalized 'review' 'PW review normalized'

    $badMembers = @{
        sheet_index_rows = @()
        sheet_documents_rows = @(
            @{
                document_guid = '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4'
                document_name = '080J082001ab001-rev.pdf'
                document_role = 'qc_pdf'
            }
        )
        sheet_packages_rows = @(@{ sheet_package_id = 'pkg-1' })
    }
    function _QDM-LoadSheetPackageQcPdfRows {
        param([string[]]$PackageIds)
        return @(
            @{
                sheet_package_id = 'pkg-1'
                qc_process_type = 'check'
                document_guid = '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4'
                document_name = '080J082001ab001-rev.pdf'
                is_active = $true
            }
        )
    }
    $bad = _QDM-BuildQcProcessTypeDiagnostics -Lookup $lookup -Members $badMembers -ProjectWiseAvailable $false
    Assert-True ($bad.checks_failed -gt 0) 'lane registry vs filename mismatch fails'
    Assert-True (@($bad.checks | Where-Object { $_.code -eq 'lane_registry_vs_filename' -and -not $_.passed }).Count -ge 1) 'reports lane_registry_vs_filename'
}

Write-Host 'test_qc_debug_mcp_process_type.ps1: OK' -ForegroundColor Green
