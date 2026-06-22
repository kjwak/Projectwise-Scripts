$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $root 'modules\Processing\QC.StatusSet.psm1') -Force

$qpdfExe = Join-Path $root 'tools\qpdf\bin\qpdf.exe'
if (-not (Test-Path -LiteralPath $qpdfExe)) {
    Write-Host "SKIP: qpdf.exe not found at $qpdfExe" -ForegroundColor Yellow
    exit 0
}

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

function _PageCount([string]$Pdf) {
    $out = & $qpdfExe --show-npages $Pdf 2>&1
    return [int](($out | Select-Object -First 1) -as [string]).Trim()
}

$tmp = Join-Path $env:TEMP ("qctest_qpdf_" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    # Use real fixture PDFs from the test corpus.
    $fixtureDir = Join-Path $root 'test\Version 01'
    $fixtures = @(
        (Join-Path $fixtureDir 'f0548dv205.pdf'),
        (Join-Path $fixtureDir 'f0548dv212.pdf'),
        (Join-Path $fixtureDir 'f0548dv213.pdf')
    )
    foreach ($f in $fixtures) {
        if (-not (Test-Path -LiteralPath $f)) { Write-Host "SKIP: missing fixture $f" -ForegroundColor Yellow; exit 0 }
    }
    $a = $fixtures[0]
    $b = $fixtures[1]
    $c = $fixtures[2]

    Write-Host "Test: single-file invokes Copy-Item shortcut and returns success" -ForegroundColor Cyan
    $out1 = Join-Path $tmp 'merged_single.pdf'
    $r1 = Merge-StatusSetPdfWithQpdf -OrderedInputPdfPaths @($a) -OutputPdf $out1 -QpdfExe $qpdfExe
    _Assert ($r1.IsSuccess) "single-file merge succeeds"
    _Assert (Test-Path -LiteralPath $out1) "single-file output file exists"
    _Assert ((_PageCount $out1) -eq (_PageCount $a)) "single-file output page count matches input"

    Write-Host "Test: multi-file merge produces sum of page counts" -ForegroundColor Cyan
    $out2 = Join-Path $tmp 'merged_multi.pdf'
    $r2 = Merge-StatusSetPdfWithQpdf -OrderedInputPdfPaths @($a, $b, $c) -OutputPdf $out2 -QpdfExe $qpdfExe
    _Assert ($r2.IsSuccess) "multi-file merge succeeds"
    _Assert (Test-Path -LiteralPath $out2) "multi-file output file exists"
    $expected = (_PageCount $a) + (_PageCount $b) + (_PageCount $c)
    $actual   = _PageCount $out2
    _Assert ($actual -eq $expected) ("merged page count $actual == sum $expected")

    Write-Host "Test: missing input returns STATUS_SET_MERGE_INPUT_MISSING (no qpdf invocation)" -ForegroundColor Cyan
    $out3 = Join-Path $tmp 'merged_missing.pdf'
    $bogus = Join-Path $tmp 'does_not_exist.pdf'
    $r3 = Merge-StatusSetPdfWithQpdf -OrderedInputPdfPaths @($a, $bogus) -OutputPdf $out3 -QpdfExe $qpdfExe
    _Assert (-not $r3.IsSuccess) "missing input returns failure"
    _Assert ($r3.Code -eq 'STATUS_SET_MERGE_INPUT_MISSING') "code is STATUS_SET_MERGE_INPUT_MISSING"

    Write-Host "Test: missing qpdf.exe returns STATUS_SET_QPDF_MISSING" -ForegroundColor Cyan
    $r4 = Merge-StatusSetPdfWithQpdf -OrderedInputPdfPaths @($a, $b) -OutputPdf $out3 -QpdfExe (Join-Path $tmp 'no_qpdf.exe')
    _Assert (-not $r4.IsSuccess) "missing qpdf returns failure"
    _Assert ($r4.Code -eq 'STATUS_SET_QPDF_MISSING') "code is STATUS_SET_QPDF_MISSING"
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failures -gt 0) { Write-Host "FAILED ($failures)" -ForegroundColor Red; exit 1 }
Write-Host "PASSED" -ForegroundColor Green
