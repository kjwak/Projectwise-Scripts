<#
.SYNOPSIS
Regression: Export-StatusSetPdfToFolder must sleep ~400 ms after each PW export
(matches legacy combine_status_set.ps1 line ~828) so Fortinet can scan the file
before the next export starts. Configurable via statusSet.pwExportThrottleMs.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Paths.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.StatusSet.psm1') -Force -DisableNameChecking | Out-Null

function Assert-True($Cond, $Msg) {
    if (-not $Cond) { throw "ASSERT FAILED: $Msg" }
}

# Stub Export-PWDocumentsSimple in the QC.StatusSet module scope so the export
# call inside Export-StatusSetPdfToFolder resolves to our fake. Write a tiny PDF
# to TargetFolder so the post-export Test-Path check passes.
$mod = Get-Module QC.StatusSet
& $mod {
    function script:Export-PWDocumentsSimple {
        [CmdletBinding()]
        param([Parameter(Mandatory)] $InputDocuments, [Parameter(Mandatory)][string]$TargetFolder)
        $name = if ($InputDocuments -is [hashtable] -and $InputDocuments.ContainsKey('Name')) { [string]$InputDocuments.Name } else { 'fake.pdf' }
        $p = Join-Path $TargetFolder $name
        Set-Content -LiteralPath $p -Value '%PDF-1.4' -Encoding ascii
    }
    function script:_SSS-PWGetDocName {
        param($Doc)
        if ($Doc -is [hashtable] -and $Doc.ContainsKey('Name')) { return [string]$Doc['Name'] }
        return 'fake.pdf'
    }
}

$tmp = Join-Path $env:TEMP ("qc_test_pwexport_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

try {
    $doc = @{ Name = 'sheet01.pdf' }
    $applyCfg = { param($Cfg) & $mod { param($C) _SSS-ApplyFsThrottleConfig -Config $C } $Cfg }

    # 1) Default throttle: ~400 ms.
    & $applyCfg @{}
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $r = Export-StatusSetPdfToFolder -InputDocument $doc -TargetFolder $tmp
    $sw.Stop()
    Assert-True $r.IsSuccess "Export should succeed (got Code='$($r.Code)' Message='$($r.Message)')"
    Assert-True ($sw.ElapsedMilliseconds -ge 350) ("Default throttle should be >=350 ms; got $($sw.ElapsedMilliseconds) ms")
    Assert-True ($sw.ElapsedMilliseconds -lt 1500) ("Default throttle should be <1500 ms; got $($sw.ElapsedMilliseconds) ms")

    # 2) Disabled via config: should be near-zero.
    & $applyCfg @{ statusSet = @{ pwExportThrottleMs = 0 } }
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $r2 = Export-StatusSetPdfToFolder -InputDocument $doc -TargetFolder $tmp
    $sw2.Stop()
    Assert-True $r2.IsSuccess "Export(disabled) should succeed"
    Assert-True ($sw2.ElapsedMilliseconds -lt 200) ("Disabled throttle should be <200 ms; got $($sw2.ElapsedMilliseconds) ms")

    # 3) Custom value: 700 ms.
    & $applyCfg @{ statusSet = @{ pwExportThrottleMs = 700 } }
    $sw3 = [System.Diagnostics.Stopwatch]::StartNew()
    $r3 = Export-StatusSetPdfToFolder -InputDocument $doc -TargetFolder $tmp
    $sw3.Stop()
    Assert-True $r3.IsSuccess "Export(700ms) should succeed"
    Assert-True ($sw3.ElapsedMilliseconds -ge 650) ("Custom 700ms throttle should be >=650 ms; got $($sw3.ElapsedMilliseconds) ms")

    Write-Host "test_pw_export_throttle: PASS" -ForegroundColor Green
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
