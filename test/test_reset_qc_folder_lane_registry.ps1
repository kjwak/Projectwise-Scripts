# Reset-QCFolderWorkflow includes lane PDF registry purge for manual delete + reset workflow.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'scripts\Reset-QCFolderWorkflow.ps1'

function Assert-True($v, $msg) { if (-not $v) { throw "ASSERT FAILED: $msg" } }

$text = Get-Content -LiteralPath $scriptPath -Raw
Assert-True ($text -match 'KeepLanePdfRegistry') 'KeepLanePdfRegistry switch documented'
Assert-True ($text -match '_RQCF-GetLanePdfNameClause') 'lane PDF name helper present'
Assert-True ($text -match 'sheet_index_lane') 'deletes lane sheet_index rows'
Assert-True ($text -match 'sheet_documents_qc_pdf') 'deletes sheet_documents qc_pdf role rows'
Assert-True ($text -match "NOT \(\`$siLaneNameClause\)") 'stem/DGN index update excludes lane rows when purging'
Assert-True ($text -match 'sheet_package_qc_pdfs') 'still clears sheet_package_qc_pdfs'

Write-Host 'OK: Reset-QCFolderWorkflow lane registry tests passed.' -ForegroundColor Green
