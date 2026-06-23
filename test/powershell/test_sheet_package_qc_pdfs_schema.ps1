# Schema v1.19: sheet_package_qc_pdfs table and reporting views.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }

$dbText = Get-Content -LiteralPath (Resolve-ModuleImplPath -ModuleName 'Core.Database.psm1') -Raw

Assert-True ($dbText -match '_QDB-GetSchemaV1dot19Additive') 'schema v1.19 additive patch exists'
Assert-True ($dbText -match "targetVersion = '1.19.0'") 'schema target version is 1.19.0'
Assert-True ($dbText -match 'CREATE TABLE sheet_package_qc_pdfs') 'sheet_package_qc_pdfs table defined'
Assert-True ($dbText -match 'ux_sheet_package_qc_pdfs_package_process_active') 'unique active index defined'
Assert-True ($dbText -match 'CREATE VIEW v_sheet_package_qc_pdf_matrix') 'matrix view defined'
Assert-True ($dbText -match 'production_qc_pdf_guid') 'matrix view has production columns'
Assert-True ($dbText -match 'check_qc_pdf_guid') 'matrix view has check columns'
Assert-True ($dbText -match 'review_qc_pdf_guid') 'matrix view has review columns'
Assert-True ($dbText -match 'production_state') 'matrix view has production_state'
Assert-True ($dbText -match 'check_state') 'matrix view has check_state'
Assert-True ($dbText -match 'review_state') 'matrix view has review_state'

$v19 = ($dbText -split 'function _QDB-GetSchemaV1dot19Additive\s*\{', 2)[1] -split 'function _QDB-GetSchemaV1dot9Additive', 2 | Select-Object -First 1
Assert-True ($v19 -match 'production_completed') 'status view exposes production_completed'
Assert-True ($v19 -match 'check_completed') 'status view exposes check_completed'
Assert-True ($v19 -match 'review_completed') 'status view exposes review_completed'
Assert-True ($v19 -match 'LEFT JOIN v_sheet_package_qc_pdf_matrix') 'status view joins lane matrix'
Assert-True ($v19 -notmatch 'sp\.pw_state_name') 'status view does not expose package pw_state_name'
Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot 'scripts\sql\backfill-sheet-package-qc-pdfs.sql')) 'backfill script exists'

Assert-True ($dbText -match '@qcLane = ''production''') 'qc_pdf_guid only updated for production lane'
Assert-True ($dbText -match 'qc_chk_pdf_guid = CASE WHEN @documentRole = ''qc_pdf'' AND @qcLane = ''check''') 'check lane uses qc_chk_pdf_guid'
Assert-True ($dbText -match 'pw_state_name = tgt.pw_state_name') 'Ensure-SheetPackage does not write package pw_state_name'
Assert-True ($dbText -match 'function Upsert-SheetPackageQcPdf') 'Upsert-SheetPackageQcPdf exported'
Assert-True ($dbText -match 'function Sync-SheetPackageLaneQcPdfsFromMembers') 'Sync-SheetPackageLaneQcPdfsFromMembers exported'
Assert-True ($dbText -match 'function Update-SheetPackageQcPdfLaneState') 'Update-SheetPackageQcPdfLaneState exported'

Write-Host 'test_sheet_package_qc_pdfs_schema: OK' -ForegroundColor Green
