# Unit tests: status-set CAD pairing stems (.dgn/.dwg/extensionless) + local folder pairing.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')
Import-Module (Resolve-ModuleImplPath -ModuleName 'Core.Results.psm1' -ModulesDir (Join-Path $root 'modules')) -Force
Import-Module (Resolve-ModuleImplPath -ModuleName 'Core.Paths.psm1' -ModulesDir (Join-Path $root 'modules')) -Force
Import-Module (Resolve-ModuleImplPath -ModuleName 'QC.StatusSet.psm1' -ModulesDir (Join-Path $root 'modules')) -Force

$failures = 0
function Assert-True([bool]$ok, [string]$msg) {
    if (-not $ok) {
        Write-Host "  FAIL: $msg" -ForegroundColor Red
        $script:failures++
    } else {
        Write-Host "  ok:   $msg" -ForegroundColor Green
    }
}

Write-Host 'Test: _SSS-TryGetCadStem / legacy paired select' -ForegroundColor Cyan
InModuleScope -ModuleName QC.StatusSet {
    Assert-True ((_SSS-TryGetCadStem -Name 'sheet.dgn') -eq 'sheet') 'sheet.dgn -> sheet'
    Assert-True ((_SSS-TryGetCadStem -Name 'sheet.DWG') -eq 'sheet') 'sheet.DWG -> sheet'
    Assert-True ((_SSS-TryGetCadStem -Name '000_G-00.01_h0757fsh') -eq '000_g-00.01_h0757fsh') 'extensionless stem keeps dots in name'
    Assert-True ($null -eq (_SSS-TryGetCadStem -Name 'sheet.pdf')) 'pdf is not CAD'
    Assert-True ($null -eq (_SSS-TryGetCadStem -Name 'notes.xlsx')) 'xlsx is not CAD'
    Assert-True ($null -eq (_SSS-TryGetCadStem -Name 'readme.txt')) 'txt is not CAD'
    Assert-True ((_SSS-IsNonCadDocumentName -Name 'notes.docx') -eq $true) 'docx is non-CAD'

    $docsDgn = @(
        [pscustomobject]@{ Name = 'A001.pdf' },
        [pscustomobject]@{ Name = 'A001.dgn' }
    )
    $pairedDgn = @(_SSS-SelectPairedPdfDocsLegacy -AllDocs $docsDgn)
    Assert-True ($pairedDgn.Count -eq 1) 'pdf+dgn pairs'
    Assert-True ((_SSS-GetDocName $pairedDgn[0]) -eq 'A001.pdf') 'paired pdf is A001.pdf'

    $docsBare = @(
        [pscustomobject]@{ Name = 'A002.pdf' },
        [pscustomobject]@{ Name = 'A002' }
    )
    $pairedBare = @(_SSS-SelectPairedPdfDocsLegacy -AllDocs $docsBare)
    Assert-True ($pairedBare.Count -eq 1) 'pdf+extensionless CAD pairs'

    $docsDwg = @(
        [pscustomobject]@{ Name = 'A003.pdf' },
        [pscustomobject]@{ Name = 'A003.dwg' }
    )
    $pairedDwg = @(_SSS-SelectPairedPdfDocsLegacy -AllDocs $docsDwg)
    Assert-True ($pairedDwg.Count -eq 1) 'pdf+dwg pairs'

    $docsAlone = @(
        [pscustomobject]@{ Name = 'A004.pdf' }
    )
    $pairedAlone = @(_SSS-SelectPairedPdfDocsLegacy -AllDocs $docsAlone)
    Assert-True ($pairedAlone.Count -eq 0) 'pdf alone does not pair'

    $docsNoise = @(
        [pscustomobject]@{ Name = 'A005.pdf' },
        [pscustomobject]@{ Name = 'A005.xlsx' },
        [pscustomobject]@{ Name = 'readme.txt' }
    )
    $pairedNoise = @(_SSS-SelectPairedPdfDocsLegacy -AllDocs $docsNoise)
    Assert-True ($pairedNoise.Count -eq 0) 'office/noise names do not count as CAD'
}

Write-Host ''
Write-Host 'Test: Get-StatusSetLocalFolderState extensionless + dwg pairing' -ForegroundColor Cyan
$tmp = Join-Path $env:TEMP ('sscadpair_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    Set-Content -LiteralPath (Join-Path $tmp 'S001.pdf') -Value 'pdf' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S001.dgn') -Value 'dgn' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S002.pdf') -Value 'pdf' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S002') -Value 'bare' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S003.pdf') -Value 'pdf' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S003.dwg') -Value 'dwg' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S004.pdf') -Value 'pdf' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S005.pdf') -Value 'pdf' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'S005.xlsx') -Value 'xls' -Encoding ascii

    $state = Get-StatusSetLocalFolderState -RootFolder $tmp
    Assert-True ($state.pairedCount -eq 3) "local pairedCount=3 (dgn+bare+dwg), got $($state.pairedCount)"
    $stems = @($state.pairedSheets | ForEach-Object { $_.stem } | Sort-Object)
    Assert-True (($stems -join ',') -eq 's001,s002,s003') "paired stems s001,s002,s003 got $($stems -join ',')"
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failures -gt 0) {
    Write-Host "FAILED: $failures assertion(s)" -ForegroundColor Red
    exit 1
}
Write-Host 'OK: status set CAD pairing tests passed.' -ForegroundColor Green
