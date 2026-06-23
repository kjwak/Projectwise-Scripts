$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $root 'modules\Processing\QC.StatusSet.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

$tmp = Join-Path $env:TEMP ("qctest_ws_" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    Write-Host "Test: Get-StatusSetWorkspaceDirectory derives a deterministic per-folder workspace key" -ForegroundColor Cyan

    $p1 = 'Documents\AZDOT 2024\AZFWY1704-FD02-SR202\CADD\Sheets'
    $p2 = 'documents/azdot 2024/azfwy1704-fd02-sr202/cadd/sheets/' # equivalent after normalization
    $p3 = 'Documents\AZDOT 2024\AZFWY2302-031_SR98\CADD\Sheets'    # different folder

    $r1 = Get-StatusSetWorkspaceDirectory -LocalRoot $tmp -SheetsFolderPath $p1
    $r2 = Get-StatusSetWorkspaceDirectory -LocalRoot $tmp -SheetsFolderPath $p2
    $r3 = Get-StatusSetWorkspaceDirectory -LocalRoot $tmp -SheetsFolderPath $p3

    _Assert ($r1.IsSuccess) "p1 succeeds"
    _Assert ($r2.IsSuccess) "p2 succeeds"
    _Assert ($r3.IsSuccess) "p3 succeeds"
    _Assert ($r1.Data.workspaceDir -eq $r2.Data.workspaceDir) "p1 and p2 normalize to the same workspaceDir"
    _Assert ($r1.Data.workspaceDir -ne $r3.Data.workspaceDir) "p1 and p3 produce different workspaceDirs"
    _Assert ($r1.Data.folderKey -and $r1.Data.folderKey.Length -eq 16) "folderKey is 16 hex chars"

    # Note: empty/null LocalRoot is rejected by [Parameter(Mandatory)] before the body runs.
    # Same for empty SheetsFolderPath. We just confirm whitespace SheetsFolderPath produces
    # the expected normalize-failure code (the body's defensive check).
    $bad = Get-StatusSetWorkspaceDirectory -LocalRoot $tmp -SheetsFolderPath '   '
    _Assert (-not $bad.IsSuccess) "whitespace path returns failure"
    _Assert ($bad.Code -eq 'STATUS_SET_PATH_NORMALIZE_FAILED') "whitespace path code is STATUS_SET_PATH_NORMALIZE_FAILED"
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failures -gt 0) { Write-Host "FAILED ($failures)" -ForegroundColor Red; exit 1 }
Write-Host "PASSED" -ForegroundColor Green
