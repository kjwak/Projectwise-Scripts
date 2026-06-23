$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\modules\Notifications\QC.NotificationTemplates.psm1') -Force
$t = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
$t['DocumentName'] = '0818000063ea501-qc.pdf'
$t['ProjectName'] = 'CAFWY2200-I-15_ELPSE'
$t['ReviewType'] = 'Production QC'
$t['WorkflowState'] = 'Corrections Received'
$out = Expand-QCNotificationTemplate -Template '[{ReviewType}] {ProjectName} | {DocumentName} | {WorkflowState}' -Tokens $t
$expected = '[Production QC] CAFWY2200-I-15_ELPSE | 0818000063ea501-qc.pdf | Corrections Received'
if ($out -ne $expected) { throw "FAIL: got '$out'" }
Write-Host "OK: $out"
