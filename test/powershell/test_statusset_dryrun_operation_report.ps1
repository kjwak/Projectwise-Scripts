$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $root 'modules\Processing\QC.StatusSet.psm1') -Force

function Assert([bool]$Cond, [string]$Msg) {
    if (-not $Cond) { throw "ASSERT FAILED: $Msg" }
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_ss_dryrun_' + [Guid]::NewGuid().ToString('n'))
$sheets = Join-Path $tmp 'Sheets'
$localRoot = Join-Path $tmp 'statussets'
$stagingRoot = Join-Path $tmp 'staging'
New-Item -ItemType Directory -Path $sheets -Force | Out-Null
try {
    Set-Content -LiteralPath (Join-Path $sheets 'A001.dgn') -Value 'cad' -Encoding ascii -NoNewline
    Set-Content -LiteralPath (Join-Path $sheets 'A001.pdf') -Value '%PDF-1.4 dryrun-a' -Encoding ascii -NoNewline
    Set-Content -LiteralPath (Join-Path $sheets 'A002.dwg') -Value 'cad' -Encoding ascii -NoNewline
    Set-Content -LiteralPath (Join-Path $sheets 'A002.pdf') -Value '%PDF-1.4 dryrun-b' -Encoding ascii -NoNewline

    $job = @{ id='dryrun-report'; type='STATUS_SET_GEN'; sourceFolder=$sheets }
    $cfg = @{
        dryRun = $true
        statusSet = @{
            mode = 'native'
            localRoot = $localRoot
            stagingRoot = $stagingRoot
            incrementalMode = $true
            retentionDays = 14
            cleanupEnabled = $false
            atomicReplaceEnabled = $true
            dryRunOperationReport = $true
            writeBackToPW = $false
            qpdfExe = (Join-Path $root 'tools\qpdf\bin\qpdf.exe')
        }
    }

    $res = Invoke-StatusSetNativeJob -Job $job -Config $cfg
    Assert ($res.IsSuccess) 'dry-run status set job succeeds'
    Assert ($res.Code -eq 'STATUS_SET_DRYRUN_REBUILD') "expected dry-run rebuild code, got $($res.Code)"
    Assert ($null -ne $res.Data.operationReport) 'operation report is present'
    Assert (@($res.Data.operationReport.writes).Count -ge 2) 'dry-run reports merge and manifest writes'
    Assert (@($res.Data.operationReport.replaces).Count -eq 1) 'dry-run reports one atomic output replace'
    Assert (@($res.Data.operationReport.skips | Where-Object { $_.action -eq 'download' -and $_.reason -eq 'local-filesystem-source' }).Count -eq 2) 'dry-run reports local sheet downloads as skipped'
    Assert (-not (Test-Path -LiteralPath $localRoot)) 'dry-run does not create localRoot files'
    Assert (-not (Test-Path -LiteralPath $stagingRoot)) 'dry-run does not create staging files when cleanup disabled'

    Write-Host 'test_statusset_dryrun_operation_report: PASS' -ForegroundColor Green
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
