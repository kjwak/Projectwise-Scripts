# Sheet package stem/role resolution (no SQL).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$dgn = Resolve-SheetPackageFromDocument -DocumentGuid '11111111-1111-1111-1111-111111111111' `
    -DocumentName '080J082001ab001.dgn' -FolderPath 'documents\proj\cadd\sheets'
Assert-Eq $dgn.sheetStem '080J082001ab001' 'DGN stem'
Assert-Eq $dgn.documentRole 'dgn' 'DGN role'
Assert-True $dgn.isSheetPackageMember 'DGN is package member'

$pdf = Resolve-SheetPackageFromDocument -DocumentGuid '22222222-2222-2222-2222-222222222222' `
    -DocumentName '080J082001ab001.pdf' -FolderPath 'documents\proj\cadd\sheets'
Assert-Eq $pdf.sheetStem '080J082001ab001' 'sheet PDF stem matches DGN'
Assert-Eq $pdf.documentRole 'sheet_pdf' 'sheet PDF role'

$qc = Resolve-SheetPackageFromDocument -DocumentGuid '33333333-3333-3333-3333-333333333333' `
    -DocumentName '080J082001ab001-qc.pdf' -FolderPath 'documents\proj\cadd\sheets'
Assert-Eq $qc.sheetStem '080J082001ab001' 'QC PDF stem strips -qc'
Assert-Eq $qc.documentRole 'qc_pdf' 'QC PDF role'

$other = Resolve-SheetPackageFromDocument -DocumentName 'folder_statusset.pdf' -FolderPath 'documents\proj\cadd\sheets'
Assert-Eq $other.documentRole 'sheet_pdf' 'status set pdf is sheet_pdf extension-wise'
Assert-True $other.isSheetPackageMember 'regular pdf is package member'

$manifest = Resolve-SheetPackageFromDocument -DocumentName 'notes.txt' -FolderPath 'documents\proj\cadd\sheets'
Assert-Eq $manifest.documentRole 'other' 'non-sheet extension is other'
Assert-True (-not $manifest.isSheetPackageMember) 'non-sheet file is not package member'

Write-Host 'OK: sheet package resolution tests passed.' -ForegroundColor Green
