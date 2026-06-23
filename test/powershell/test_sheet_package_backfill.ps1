# Sheet package backfill plan tests (mirrors scripts/sql/backfill-sheet-packages.sql).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$folder = 'Documents\Proj\CADD\Sheets'
$dgnGuid = '11111111-1111-1111-1111-111111111111'
$pdfGuid = '22222222-2222-2222-2222-222222222222'
$qcGuid = '33333333-3333-3333-3333-333333333333'

# 1. Three legacy rows -> one package + three documents
$rows = @(
    @{
        documentGuid = $dgnGuid; documentName = '080J082001ab001.dgn'; folderPath = $folder
        designerEmail = 'd@example.com'; pwStateName = 'QC Complete'
        productionQcCompletedCount = 2
    },
    @{
        documentGuid = $pdfGuid; documentName = '080J082001ab001.pdf'; folderPath = $folder
        qcCycleId = 'cycle-1'; qcCycleNumber = '1'; qcReviewType = 'Production QC'
    },
    @{
        documentGuid = $qcGuid; documentName = '080J082001ab001-qc.pdf'; folderPath = $folder
        pwStateName = 'QC Complete'
    }
)
$plan = Build-SheetPackageBackfillPlan -SheetIndexRows $rows
Assert-Eq $plan.packageCount 1 'three siblings should produce one package'
Assert-Eq $plan.documentCount 3 'three siblings should produce three documents'
Assert-Eq $plan.packages[0].sheetStem '080J082001ab001' 'package stem'
Assert-Eq $plan.packages[0].dgnGuid.ToString() $dgnGuid 'package dgn guid'
Assert-Eq $plan.packages[0].sheetPdfGuid.ToString() $pdfGuid 'package sheet pdf guid'
Assert-Eq $plan.packages[0].qcPdfGuid.ToString() $qcGuid 'package qc pdf guid'
Assert-Eq $plan.packages[0].qcCycleId 'cycle-1' 'cycle from sheet pdf'
Assert-Eq $plan.packages[0].productionQcCompletedCount 2 'completion rollup from dgn'

# 6. sheet_index.sheet_package_id links for all siblings
Assert-True ($plan.indexLinks.ContainsKey($dgnGuid.ToLowerInvariant())) 'dgn index link'
Assert-True ($plan.indexLinks.ContainsKey($pdfGuid.ToLowerInvariant())) 'pdf index link'
Assert-True ($plan.indexLinks.ContainsKey($qcGuid.ToLowerInvariant())) 'qc pdf index link'
Assert-Eq $plan.indexLinks[$dgnGuid.ToLowerInvariant()].ToString() $plan.packages[0].sheetPackageId.ToString() 'consistent package id'

# 2. Missing QC PDF still creates package
$planMissingQc = Build-SheetPackageBackfillPlan -SheetIndexRows @(
    @{ documentGuid = $dgnGuid; documentName = '080J082001ab002.dgn'; folderPath = $folder },
    @{ documentGuid = $pdfGuid; documentName = '080J082001ab002.pdf'; folderPath = $folder }
)
Assert-Eq $planMissingQc.packageCount 1 'pdf+dgn without qc pdf still creates package'
Assert-Eq $planMissingQc.documentCount 2 'pdf+dgn document count'
Assert-Eq $planMissingQc.packages[0].qcPdfGuid $null 'qc pdf guid remains null'

# 3. Orphan QC PDF still creates package
$orphanQcGuid = '44444444-4444-4444-4444-444444444444'
$planOrphanQc = Build-SheetPackageBackfillPlan -SheetIndexRows @(
    @{ documentGuid = $orphanQcGuid; documentName = '080J082001ab003-qc.pdf'; folderPath = $folder; pwStateName = 'QC Received' }
)
Assert-Eq $planOrphanQc.packageCount 1 'orphan qc pdf creates package'
Assert-Eq $planOrphanQc.documentCount 1 'orphan qc pdf creates one document'
Assert-Eq $planOrphanQc.packages[0].qcPdfGuid.ToString() $orphanQcGuid 'orphan qc pdf guid populated'

# Linked QC PDF from qc_pdf_guid when QC row is absent
$linkedQcGuid = '55555555-5555-5555-5555-555555555555'
$planLinkedQc = Build-SheetPackageBackfillPlan -SheetIndexRows @(
    @{
        documentGuid = $pdfGuid; documentName = '080J082001ab004.pdf'; folderPath = $folder
        qcPdfGuid = $linkedQcGuid; qcPdfName = '080J082001ab004-qc.pdf'
    },
    @{ documentGuid = $dgnGuid; documentName = '080J082001ab004.dgn'; folderPath = $folder }
)
Assert-Eq $planLinkedQc.documentCount 3 'linked qc pdf adds third document'
Assert-True (@($planLinkedQc.documents | Where-Object { $_.documentRole -eq 'qc_pdf' }).Count -eq 1) 'linked qc pdf role present'

# 4. Re-running backfill plan does not duplicate rows
$planRerun = Build-SheetPackageBackfillPlan -SheetIndexRows $rows
Assert-Eq $planRerun.packageCount $plan.packageCount 'rerun package count stable'
Assert-Eq $planRerun.documentCount $plan.documentCount 'rerun document count stable'
Assert-Eq $planRerun.packages[0].sheetPackageId.ToString() $plan.packages[0].sheetPackageId.ToString() 'deterministic package id'

# 5. Duplicate role conflict is reported
$dupDgn1 = '66666666-6666-6666-6666-666666666666'
$dupDgn2 = '77777777-7777-7777-7777-777777777777'
$planDup = Build-SheetPackageBackfillPlan -SheetIndexRows @(
    @{ documentGuid = $dupDgn1; documentName = '080J082001ab005.dgn'; folderPath = $folder },
    @{ documentGuid = $dupDgn2; documentName = '080J082001ab005.dgn'; folderPath = $folder }
)
$dupConflicts = @($planDup.conflicts | Where-Object { $_.conflictType -eq 'duplicate_role' })
Assert-True ($dupConflicts.Count -ge 2) 'duplicate dgn conflict logged for each competitor'
Assert-Eq $planDup.documentCount 0 'duplicate role documents are not inserted'
Assert-Eq $planDup.packages[0].dgnGuid $null 'ambiguous dgn guid not chosen'

# 7. Event/history rows can resolve sheet_package_id via document_guid
$historyGuid = $pdfGuid.ToLowerInvariant()
Assert-True ($plan.indexLinks.ContainsKey($historyGuid)) 'history document guid maps to package'
$eventPackageId = $plan.indexLinks[$historyGuid]
Assert-Eq $eventPackageId.ToString() $plan.packages[0].sheetPackageId.ToString() 'event backfill target package id'

# Invalid GUID conflict
$planInvalid = Build-SheetPackageBackfillPlan -SheetIndexRows @(
    @{ documentGuid = 'short'; documentName = '080J082001ab006.pdf'; folderPath = $folder }
)
Assert-True (@($planInvalid.conflicts | Where-Object { $_.conflictType -eq 'invalid_guid' }).Count -eq 1) 'invalid guid conflict'

# SQL scripts exist
$sqlBackfill = Join-Path $repoRoot 'scripts\sql\backfill-sheet-packages.sql'
$sqlValidate = Join-Path $repoRoot 'scripts\sql\validate-sheet-packages.sql'
Assert-True (Test-Path -LiteralPath $sqlBackfill) 'backfill SQL script exists'
Assert-True (Test-Path -LiteralPath $sqlValidate) 'validate SQL script exists'

Write-Host 'OK: sheet package backfill tests passed.' -ForegroundColor Green
