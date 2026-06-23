$ErrorActionPreference = 'Stop'
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}
function Assert-Null($Actual, $Message) {
    if ($null -ne $Actual) { throw "ASSERT FAILED: $Message`nExpected null, got: $Actual" }
}

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Discovery.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force

Assert-Eq (Get-PWSheetStemFromDocumentName -DocumentName 'CA001-prod.pdf') 'CA001' 'strip prod suffix'
Assert-Eq (Get-PWSheetStemFromDocumentName -DocumentName 'CA001-chk.pdf') 'CA001' 'strip chk suffix'
Assert-Eq (Get-PWSheetStemFromDocumentName -DocumentName 'CA001-rev.pdf') 'CA001' 'strip rev suffix'
Assert-Eq (Get-PWSheetStemFromDocumentName -DocumentName 'CA001-qc.pdf') 'CA001-qc' 'legacy -qc not production alias'

Assert-Eq (Get-PWQcPdfLaneFromDocumentName -DocumentName 'CA001-prod.pdf') 'production' 'lane prod'
Assert-Eq (Get-PWQcPdfLaneFromDocumentName -DocumentName 'CA001-chk.pdf') 'check' 'lane chk'
Assert-Eq (Get-PWQcPdfLaneFromDocumentName -DocumentName 'CA001-rev.pdf') 'review' 'lane rev'
Assert-Null (Get-PWQcPdfLaneFromDocumentName -DocumentName 'CA001-qc.pdf') 'legacy qc not a lane'

$syncNames = Get-PWAssociatedSheetSyncDocumentNames -SheetStem 'CA001'
Assert-True ($syncNames -contains 'CA001-prod.pdf') 'sync includes prod pdf'
Assert-True (-not ($syncNames -contains 'CA001-chk.pdf')) 'sync excludes chk pdf'

Write-Host 'test_qc_lane_pdf_resolver.ps1: OK'
