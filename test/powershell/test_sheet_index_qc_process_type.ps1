# qc_process_type column normalization and MCP folder-path parsing.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Diagnostics\QC.DebugMcp.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }

InModuleScope -ModuleName Core.Database {
    Assert-Eq (_QDB-NormalizeQcProcessTypeForColumn -RawValue 'Review') 'review' 'normalizes Review'
    Assert-Eq (_QDB-NormalizeQcProcessTypeForColumn -RawValue '') $null 'empty stays null'
}

InModuleScope -ModuleName QC.DebugMcp {
    $folder = _QDM-ParseDocumentPath -DocumentPath 'pw:\\server\Documents\Caltrans\Proj\CADD\Sheets\Seg_1\'
    Assert-True $folder.is_folder_only 'trailing slash is folder-only'
    Assert-Eq $folder.document_name $null 'no document name for folder path'
    Assert-True ($folder.folder_path -match 'seg_1') 'folder path includes seg_1'

    $doc = _QDM-ParseDocumentPath -DocumentPath 'Documents\Caltrans\Proj\CADD\Sheets\Seg_1\0818000063ea501.pdf'
    Assert-True (-not $doc.is_folder_only) 'pdf path is document lookup'
    Assert-Eq $doc.document_name '0818000063ea501.pdf' 'document name parsed'

    $folderBare = _QDM-ParseDocumentPath -DocumentPath 'Documents\Caltrans\Proj\CADD\Sheets\Seg_1'
    Assert-True $folderBare.is_folder_only 'extensionless terminal segment is folder-only'
}

Write-Host 'test_sheet_index_qc_process_type.ps1: OK' -ForegroundColor Green
