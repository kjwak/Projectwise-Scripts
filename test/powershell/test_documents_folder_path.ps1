$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "FAIL: $msg (got '$a', want '$b')" }
}

$r = Normalize-QCDocumentsFolderPath -Path 'Caltrans\CAFWY2200\CADD\Sheets\Seg_1'
Assert-Eq $r.Data.path 'documents\caltrans\cafwy2200\cadd\sheets\seg_1' 'prepend documents'

$r2 = Normalize-QCDocumentsFolderPath -Path 'Documents\Caltrans\CAFWY2200\CADD\Sheets\Seg_1'
Assert-Eq $r2.Data.path 'documents\caltrans\cafwy2200\cadd\sheets\seg_1' 'already documents'

$r3 = Normalize-QCDocumentsFolderPath -Path 'pw:\\ds\Documents\AZDOT 2024\Proj\CADD\Sheets'
Assert-Eq ($r3.Data.path.StartsWith('documents\azdot 2024')) $true 'pw strip'

Write-Host 'OK test_documents_folder_path'
