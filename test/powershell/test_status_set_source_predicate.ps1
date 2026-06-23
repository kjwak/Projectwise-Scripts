# Unit tests for Test-QCStatusSetSourceDocument (audit STATUS_SET_GEN source predicate).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Queue\QC.Triggers.psm1') -Force

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

function _Test-StatusSetSource($expect, $docName, $folderPath, $label) {
    $actual = Test-QCStatusSetSourceDocument -DocumentName $docName -FolderPath $folderPath
    _Assert ($actual -eq $expect) "$label expected $expect got $actual ($docName @ $folderPath)"
}

$seg = 'Documents\AZDOT 2024\AZFWY2302-32 US60 Schulze Ranch to Allred Rd ER\CADD\Sheets\Seg_1'
$sheets = 'Documents\AZDOT 2024\AZFWY2302-32 US60 Schulze Ranch to Allred Rd ER\CADD\Sheets'
$refOrd = 'Documents\AZDOT 2024\AZFWY2302-32 US60 Schulze Ranch to Allred Rd ER\CADD\Ref-ORD'
$refNoOrd = 'Documents\AZDOT 2024\AZFWY2302-32 US60 Schulze Ranch to Allred Rd ER\CADD\Ref-NoORD'
$working = 'Documents\AZDOT 2024\AZFWY2302-32 US60 Schulze Ranch to Allred Rd ER\CADD\Working'
$templates = 'Documents\AZDOT 2024\AZFWY2302-32 US60 Schulze Ranch to Allred Rd ER\CADD\Templates'

# Should return true
_Test-StatusSetSource $true '080J082001ab001.dgn' $seg 'dgn in Seg_1'
_Test-StatusSetSource $true '080J082001ab001.pdf' $seg 'normal pdf in Seg_1'
_Test-StatusSetSource $true '170_TF-61.02a_o0562pln-REV00-RFI00E.pdf' $sheets 'normal pdf in Sheets root'

# Filename containing "qc" but not a QC artifact suffix — must still pass
_Test-StatusSetSource $true '170_TF-61.02a_o0562pln-REV00-RFI00E.pdf' $sheets 'pdf with qc substring in stem only via REV path'

# Should return false — QC artifacts
_Test-StatusSetSource $false '080J082001ab001-qc.pdf' $seg 'dash-qc pdf'
_Test-StatusSetSource $false '0818000063ea502-qc.pdf' $seg 'dash-qc pdf 2'
_Test-StatusSetSource $false '080J082001ab001_qc.pdf' $seg 'underscore-qc pdf'
_Test-StatusSetSource $false 'A101_QC.pdf' $seg 'stem ends with _QC'
_Test-StatusSetSource $false '_StatusSet.pdf' $seg 'status set output pdf'

# Should return false — non-sheet extensions / paths
_Test-StatusSetSource $false 'F0893 US 60 Schulze to Allred_ORD Templatel.itl' $refOrd 'itl in Ref-ORD'
_Test-StatusSetSource $false 'Road_Civil Templates_Seg 2.itl' $templates 'itl in Templates'
_Test-StatusSetSource $false 'e1eb559eef1b9bb9_2265_001_of_001.PwPerfDoNotUse' $seg 'PwPerfDoNotUse'
_Test-StatusSetSource $false 'report.xlsx' $seg 'xlsx'
_Test-StatusSetSource $false 'deck.pptx' $seg 'pptx'

# Should return false — DGN/PDF under excluded CADD subtrees
_Test-StatusSetSource $false '080J082001ab001.dgn' $working 'dgn in Working'
_Test-StatusSetSource $false '080J082001ab001.pdf' $working 'pdf in Working'
_Test-StatusSetSource $false '080J082001ab001.dgn' $refOrd 'dgn in Ref-ORD'
_Test-StatusSetSource $false '080J082001ab001.pdf' $refOrd 'pdf in Ref-ORD'
_Test-StatusSetSource $false '080J082001ab001.dgn' $refNoOrd 'dgn in Ref-NoORD'
_Test-StatusSetSource $false '080J082001ab001.pdf' $refNoOrd 'pdf in Ref-NoORD'

Write-Host 'OK: status set source predicate tests passed.' -ForegroundColor Green
