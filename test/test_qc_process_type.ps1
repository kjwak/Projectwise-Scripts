$ErrorActionPreference = 'Stop'
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}
function Assert-Null($Actual, $Message) {
    if ($null -ne $Actual) { throw "ASSERT FAILED: $Message`nExpected null, got: $Actual" }
}
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force

Assert-Eq (Normalize-QCProcessType -ProcessType 'Production QC') 'production' 'Production QC -> production'
Assert-Eq (Normalize-QCProcessType -ProcessType 'production') 'production' 'production -> production'
Assert-Eq (Normalize-QCProcessType -ProcessType 'Independent Check') 'check' 'Independent Check -> check'
Assert-Eq (Normalize-QCProcessType -ProcessType 'independent_check') 'check' 'independent_check -> check'
Assert-Eq (Normalize-QCProcessType -ProcessType 'Peer Review') 'review' 'Peer Review -> review'
Assert-Eq (Normalize-QCProcessType -ProcessType 'peer_review') 'review' 'peer_review -> review'
Assert-Null (Normalize-QCProcessType -ProcessType 'Unknown Type') 'Unknown fails'

Assert-Eq (Get-QCProcessTypePdfSuffix -ProcessType 'production') 'prod' 'production suffix'
Assert-Eq (Get-QCProcessTypePdfSuffix -ProcessType 'check') 'chk' 'check suffix'
Assert-Eq (Get-QCProcessTypePdfSuffix -ProcessType 'review') 'rev' 'review suffix'
Assert-Eq (Get-QCLaneQcPdfExpectedName -SheetBaseName 'CA001' -ProcessType 'production') 'CA001-prod.pdf' 'prod pdf name'
Assert-Eq (Get-QCLaneQcPdfExpectedName -SheetBaseName 'CA001' -ProcessType 'check') 'CA001-chk.pdf' 'chk pdf name'

Assert-False (Test-QCLegacySiblingStateSyncEnabled -Config @{}) 'legacy sibling sync off by default'
Assert-False (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType 'production' -Config @{}) 'production sibling sync off by default'

Write-Host 'test_qc_process_type.ps1: OK'
