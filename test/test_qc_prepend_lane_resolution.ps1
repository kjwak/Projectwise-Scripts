# QC_PREPEND lane resolution: process type -> *-prod/*-chk/*-rev.pdf before prepend execution.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Processors.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Notifications.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-True($c, $msg) { if (-not $c) { throw "ASSERT FAILED: $msg" } }
function Assert-Match($h, $p, $msg) { if ($h -notmatch $p) { throw "ASSERT FAILED: $msg (haystack='$h')" } }

InModuleScope -ModuleName QC.Processors {
    function Get-PWQcPrependRoleFieldsFromSourcePdf {
        param($FolderPath, $SourceDocumentName, $Config)
        return @{ found = $false; qcProcessType = ''; qcReviewType = '' }
    }
    function Get-PWQcPdfLaneFromDocumentName {
        param([string]$DocumentName)
        if ($DocumentName -match '(?i)-chk\.pdf$') { return 'check' }
        if ($DocumentName -match '(?i)-rev\.pdf$') { return 'review' }
        if ($DocumentName -match '(?i)-prod\.pdf$') { return 'production' }
        return $null
    }
    function Invoke-QCDatabaseQuery {
        param($Config, $Sql, $Parameters = @{})
        $table = New-Object System.Data.DataTable
        [void]$table.Columns.Add('qc_process_type', [string])
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ table = $table }
    }
    function Test-QCDatabaseEnabled { param($Config) return $true }

    $cfg = @{ database = @{ enabled = $true }; qcPrepend = @{ historyRoot = 'C:\hist' } }
    $stemJob = @{
        id = 'j-review'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = '080J082001ab001.pdf'
        metadata = @{ qcProcessType = 'review' }
    }
    $lane = _QCP-TryResolvePrependLaneContext -Job $stemJob -Config $cfg
    Assert-True $lane.IsSuccess 'review metadata resolves'
    Assert-Eq $lane.Data.qcProcessType 'review' 'process type review'
    Assert-Eq $lane.Data.pdfSuffix 'rev' 'suffix rev'
    Assert-Eq $lane.Data.expectedLanePdfName '080J082001ab001-rev.pdf' 'expected *-rev.pdf'
    Assert-Eq $lane.Data.resolutionSource 'job_metadata' 'metadata wins when qcProcessType present'

    $missing = @{
        id = 'j-missing'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = '080J082001ab001.pdf'
        metadata = @{}
    }
    $fail = _QCP-TryResolvePrependLaneContext -Job $missing -Config $cfg
    Assert-True (-not $fail.IsSuccess) 'missing process type fails'
    Assert-Eq $fail.Code 'QC_PROCESS_TYPE_UNKNOWN' 'unknown process type code'

    foreach ($pair in @(
            @{ type = 'check'; suffix = 'chk'; name = '080J082001ab001-chk.pdf' },
            @{ type = 'production'; suffix = 'prod'; name = '080J082001ab001-prod.pdf' }
        )) {
        $job = @{
            id = ('j-' + $pair.type)
            sourceFolder = 'Documents\X\CADD\Sheets'
            sourceName = '080J082001ab001.pdf'
            metadata = @{ qcProcessType = $pair.type }
        }
        $res = _QCP-TryResolvePrependLaneContext -Job $job -Config $cfg
        Assert-True $res.IsSuccess ($pair.type + ' resolves')
        Assert-Eq $res.Data.expectedLanePdfName $pair.name ($pair.type + ' lane pdf name')
        Assert-True ($res.Data.expectedLanePdfName -notmatch '-prod\.pdf$' -or $pair.type -eq 'production') ($pair.type + ' must not use wrong suffix')
    }
}

# legacy/prepend_qc.ps1 honors lane parameters (no PW connect; exits after param resolution logging would need mock - use strict param block)
$legacyScript = Join-Path $repoRoot 'legacy\prepend_qc.ps1'
Assert-True (Test-Path -LiteralPath $legacyScript) 'legacy prepend script exists'
$legacyText = Get-Content -LiteralPath $legacyScript -Raw
Assert-Match $legacyText 'QcProcessType' 'legacy script defines QcProcessType'
Assert-Match $legacyText 'QcPdfSuffix' 'legacy script defines QcPdfSuffix'
Assert-Match $legacyText 'HistoryDocumentName' 'legacy script defines HistoryDocumentName'
Assert-Match $legacyText 'QC_PROCESS_TYPE_UNKNOWN' 'legacy script fails without lane in strict mode'

$settings = Get-QCNotificationSettings -Config @{}
$jobObj = [pscustomobject]@{ metadata = [pscustomobject]@{ qcProcessType = 'review' } }
$jobHt = @{ metadata = @{ qcProcessType = 'review' } }
$event = @{
    sheetStem = '080J082001ab001'
    previousState = 'Initiate Origination'
    currentState = 'In Development'
    transitionSource = 'user_audit'
    logicalTransitionAnchor = 'audit:1'
    recipientKey = 'r@example.com'
}
$dedupeFromJob = Get-QCNotificationDedupeKey -Event ($event.Clone()) -Settings $settings -Config @{} -Job $jobHt
Assert-Match $dedupeFromJob 'qcProcessType=review' 'dedupe key includes normalized qcProcessType from job metadata'

Write-Host 'test_qc_prepend_lane_resolution.ps1: OK' -ForegroundColor Green
