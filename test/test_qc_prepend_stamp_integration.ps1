<#
.SYNOPSIS
Integration tests: resolve I15 review/check stamps from appsettings and apply via overlay exe.
#>
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force

$settingsPath = Join-Path $repoRoot 'appsettings.json'
$read = Read-QCAppSettings -Path $settingsPath
Assert-True $read.IsSuccess 'appsettings.json loads'
$config = $read.Data.config

Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ReviewStamp.psm1') -Force -DisableNameChecking

$i15Folder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$stemPdf = '0818000063ea502.pdf'

foreach ($processType in @('review', 'check')) {
    $resolved = Resolve-QCStampForProcess -Config $config -ProcessType $processType -FolderPath $i15Folder
    Assert-True $resolved.IsSuccess "I15 $processType stamp resolves from appsettings"
    Assert-Eq $resolved.resolvedStampProfile 'I15_ELPSE' "I15 $processType uses I15_ELPSE profile"
    Assert-Eq $resolved.resolvedStampName 'I15_DR' "I15 $processType uses I15_DR asset"
    Assert-True (Test-Path -LiteralPath $resolved.stampPath) "I15_DR stamp file exists: $($resolved.stampPath)"
}

$reviewResolved = Resolve-QCStampForProcess -Config $config -ProcessType 'review' -FolderPath $i15Folder
$layout = Get-QCStampProfileLayout -Config $config -StampProfile 'I15_ELPSE'
Assert-Eq $layout.stampHeightPt 300 'I15 profile layout height from appsettings'

$overlayExe = Resolve-QCReviewStampOverlayExe -PreferredPath ([string]$config.qcPrepend.overlayExePath) -RepoRoot $repoRoot
Assert-True (Test-Path -LiteralPath $overlayExe) "overlay exe exists: $overlayExe"

$sourcePdf = Join-Path $repoRoot 'test\output\f0548dv206_qc_v00_to_v01.pdf'
if (-not (Test-Path -LiteralPath $sourcePdf)) {
    throw "Fixture PDF missing: $sourcePdf"
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("qc_stamp_test_" + [Guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
$stampedPdf = Join-Path $tempDir 'stamped-rev.pdf'
Copy-Item -LiteralPath $sourcePdf -Destination $stampedPdf -Force
$beforeSize = (Get-Item -LiteralPath $stampedPdf).Length

$stampParams = @{
    OverlayExe = $overlayExe
    PdfPath = $stampedPdf
    StampPath = [string]$reviewResolved.stampPath
    StampHeightPt = [double]$layout.stampHeightPt
    MarginOutsidePt = if ($null -ne $layout.marginOutsidePt) { [double]$layout.marginOutsidePt } else { 12 }
    PopulateTextFields = $false
}
if ($null -ne $layout.stampXPt -and $null -ne $layout.stampYPt) {
    $stampParams['StampXPt'] = [double]$layout.stampXPt
    $stampParams['StampYPt'] = [double]$layout.stampYPt
}

$result = Invoke-QCReviewStamp @stampParams
if (-not $result.applied) {
    $detail = if ($result.reason) { $result.reason } else { 'unknown' }
    if ($result.stdout) { $detail += "`n$($result.stdout)" }
    throw "ASSERT FAILED: overlay stamp apply failed for I15 review`n$detail"
}

$afterSize = (Get-Item -LiteralPath $stampedPdf).Length
Assert-True ($afterSize -gt 0) 'stamped PDF exists'
Assert-True ($afterSize -ne $beforeSize) 'stamped PDF content changed after overlay apply'

# Lane suffix inference (prepend script logic)
$historyName = '0818000063ea502-rev.pdf'
$inferred = ''
if ([string]$historyName -match '(?i)-(prod|chk|rev)\.pdf$') {
    switch ($Matches[1].ToLowerInvariant()) {
        'prod' { $inferred = 'production' }
        'chk' { $inferred = 'check' }
        'rev' { $inferred = 'review' }
    }
}
Assert-Eq $inferred 'review' 'lane PDF suffix infers review process type'

Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host 'test_qc_prepend_stamp_integration.ps1: OK'
