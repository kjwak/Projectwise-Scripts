<#
.SYNOPSIS
Local offline test: apply a configured stamp profile to a PDF (copies stamp markup like Bluebeam paste).

Does not connect to ProjectWise, SQL, or the QC queue.

.EXAMPLE
  .\test_qc_prepend_stamp_integration.ps1

.EXAMPLE
  .\test_qc_prepend_stamp_integration.ps1 -StampProfile I15_ELPSE -ProcessType check
#>
[CmdletBinding()]
param(
    [string]$TargetPdf = '',
    [string]$OutputPdf = '',
    [string]$AppsettingsPath = '',
    [string]$StampProfile = 'I15_ELPSE',
    [ValidateSet('check', 'review', 'production')]
    [string]$ProcessType = 'check',
    [string]$StampAssetName = '',
    [switch]$PopulateTextFields,
    [switch]$FlattenStampAnnotation,
    [string]$Label = ''
)

$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Get-ConfigSection {
    param([object]$Root, [string]$Name)
    if (-not $Root) { return $null }
    if ($Root -is [hashtable]) {
        if ($Root.ContainsKey($Name)) { return $Root[$Name] }
        foreach ($k in $Root.Keys) {
            if ([string]$k -eq $Name) { return $Root[$k] }
        }
        return $null
    }
    if ($Root.PSObject.Properties[$Name]) { return $Root.PSObject.Properties[$Name].Value }
    return $null
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.ReviewStamp.psm1') -Force -DisableNameChecking

if ([string]::IsNullOrWhiteSpace($AppsettingsPath)) {
    $AppsettingsPath = Join-Path $repoRoot 'appsettings.json'
} elseif (-not [System.IO.Path]::IsPathRooted($AppsettingsPath)) {
    $AppsettingsPath = Join-Path $repoRoot $AppsettingsPath
}
$read = Read-QCAppSettings -Path $AppsettingsPath
Assert-True $read.IsSuccess "appsettings load failed: $($read.Message)"
$config = $read.Data.config

$qcProcess = Get-ConfigSection -Root $config -Name 'QCProcess'
if (-not $qcProcess) { throw 'appsettings.json missing QCProcess section.' }

$profiles = Get-ConfigSection -Root $qcProcess -Name 'StampProfiles'
$profile = Get-ConfigSection -Root $profiles -Name $StampProfile
if (-not $profile) { throw "Stamp profile '$StampProfile' not found under QCProcess.StampProfiles." }

if ([string]::IsNullOrWhiteSpace($StampAssetName)) {
    $StampAssetName = [string](Get-ConfigSection -Root $profile -Name $ProcessType)
}
if ([string]::IsNullOrWhiteSpace($StampAssetName)) {
    throw "Profile '$StampProfile' has no stamp asset for process type '$ProcessType'."
}

$stampAssets = Get-ConfigSection -Root $qcProcess -Name 'StampAssets'
$stampRel = [string](Get-ConfigSection -Root $stampAssets -Name $StampAssetName)
if ([string]::IsNullOrWhiteSpace($stampRel)) {
    throw "Stamp asset '$StampAssetName' not found under QCProcess.StampAssets."
}
$stampPath = if ([System.IO.Path]::IsPathRooted($stampRel)) { $stampRel } else { Join-Path $repoRoot $stampRel }
Assert-True (Test-Path -LiteralPath $stampPath) "stamp PDF missing: $stampPath"

$stampHeightPt = 300
$marginOutsidePt = 12
$stampXPt = $null
$stampYPt = $null
$populateFromProfile = $false
$layout = Get-ConfigSection -Root $profile -Name 'layout'
if ($layout) {
    if ($null -ne $layout.stampHeightPt) { $stampHeightPt = [double]$layout.stampHeightPt }
    if ($null -ne $layout.marginOutsidePt) { $marginOutsidePt = [double]$layout.marginOutsidePt }
    if ($null -ne $layout.populateTextFields) { $populateFromProfile = [bool]$layout.populateTextFields }
    $pos = Get-ConfigSection -Root $layout -Name 'stampPositionPt'
    if ($pos) {
        if ($null -ne $pos.x) { $stampXPt = [double]$pos.x }
        if ($null -ne $pos.y) { $stampYPt = [double]$pos.y }
    }
}
$usePopulateTextFields = if ($PopulateTextFields.IsPresent) { $true } else { $populateFromProfile }

$stampPy = Join-Path $repoRoot 'overlay\qc_review_stamp.py'
Assert-True (Test-Path -LiteralPath $stampPy) 'overlay/qc_review_stamp.py missing'

if ([string]::IsNullOrWhiteSpace($TargetPdf)) {
    $TargetPdf = Join-Path $repoRoot 'test\050_D-02.10_d0847drn-qc.pdf'
} elseif (-not [System.IO.Path]::IsPathRooted($TargetPdf)) {
    $TargetPdf = Join-Path $repoRoot $TargetPdf
}
$TargetPdf = [System.IO.Path]::GetFullPath($TargetPdf)
Assert-True (Test-Path -LiteralPath $TargetPdf) "target PDF missing: $TargetPdf"

if ([string]::IsNullOrWhiteSpace($Label)) {
    $mode = if ($FlattenStampAnnotation.IsPresent) { 'flattened' } else { 'copy' }
    $Label = ($StampProfile.ToLowerInvariant() + '-' + $ProcessType + '-' + $mode)
}
if ([string]::IsNullOrWhiteSpace($OutputPdf)) {
    $outDir = Join-Path $repoRoot 'test\output'
    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($TargetPdf)
    $OutputPdf = Join-Path $outDir ($base + '-' + $Label + '-stamped.pdf')
} elseif (-not [System.IO.Path]::IsPathRooted($OutputPdf)) {
    $OutputPdf = Join-Path $repoRoot $OutputPdf
}
$OutputPdf = [System.IO.Path]::GetFullPath($OutputPdf)
$outParent = Split-Path -Parent $OutputPdf
if (-not (Test-Path -LiteralPath $outParent)) {
    New-Item -ItemType Directory -Path $outParent -Force | Out-Null
}

Copy-Item -LiteralPath $TargetPdf -Destination $OutputPdf -Force
$beforeSize = (Get-Item -LiteralPath $OutputPdf).Length

$pyArgs = @(
    $stampPy,
    $OutputPdf,
    $stampPath,
    '--stamp-height-pt', [string]$stampHeightPt
)
if ($null -ne $stampXPt -and $null -ne $stampYPt) {
    $pyArgs += @('--stamp-x-pt', [string]$stampXPt, '--stamp-y-pt', [string]$stampYPt)
} else {
    $pyArgs += @('--margin-inset-pt', [string]$marginOutsidePt)
}
if (-not $usePopulateTextFields) {
    $pyArgs += '--no-populate-text-fields'
}
if ($FlattenStampAnnotation.IsPresent) {
    $pyArgs += '--flatten-stamp-annotation'
}

Write-Host "Running stamp: $($pyArgs -join ' ')"
& python @pyArgs
if ($LASTEXITCODE -ne 0) {
    Remove-Item -LiteralPath $OutputPdf -Force -ErrorAction SilentlyContinue
    throw "python qc_review_stamp.py failed with exit code $LASTEXITCODE"
}

$afterSize = (Get-Item -LiteralPath $OutputPdf).Length
Assert-True ($afterSize -gt 0) 'stamped PDF is empty'
Assert-True ($afterSize -ne $beforeSize) 'stamped PDF unchanged (stamp may have failed silently)'

Write-Host 'Local stamp test passed.' -ForegroundColor Green
Write-Host "  profile:     $StampProfile ($ProcessType -> $StampAssetName)"
Write-Host "  source:      $TargetPdf"
Write-Host "  output:      $OutputPdf"
Write-Host "  stamp asset: $stampPath"
Write-Host "  mode:        $(if ($FlattenStampAnnotation.IsPresent) { 'flattened Stamp annotation' } else { 'copy markup (default)' })"
Write-Host "  populate:    $usePopulateTextFields"
