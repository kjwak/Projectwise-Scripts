# Update-SheetPackageQcPdfLaneState must scope by qc_process_type (and optional sheet_package_id).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

$dbText = Get-Content -LiteralPath (Resolve-ModuleImplPath -ModuleName 'Core.Database.psm1') -Raw
if ($dbText -notmatch 'function Update-SheetPackageQcPdfLaneState') {
    throw 'Update-SheetPackageQcPdfLaneState not found'
}
$fn = ($dbText -split 'function Update-SheetPackageQcPdfLaneState', 2)[1] -split 'function Sync-SheetPackageLaneQcPdfsFromMembers', 2 | Select-Object -First 1
if ($fn -notmatch 'qc_process_type = @qcProcessType') {
    throw 'ASSERT FAILED: lane state update must filter by qc_process_type'
}
if ($fn -notmatch 'AND is_active = 1') {
    throw 'ASSERT FAILED: lane state update must target active rows only'
}
if ($fn -notmatch 'sheet_package_id = @sheetPackageId') {
    throw 'ASSERT FAILED: lane state update should support sheet_package_id scoping'
}

Write-Host 'test_update_sheet_package_qc_pdf_lane_state: OK' -ForegroundColor Green
