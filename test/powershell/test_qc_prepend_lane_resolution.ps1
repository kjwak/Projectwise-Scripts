# QC_PREPEND lane resolution: process type -> *-prod/*-chk/*-rev.pdf before prepend execution.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.Processors.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.Notifications.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-True($c, $msg) { if (-not $c) { throw "ASSERT FAILED: $msg" } }
function Assert-Match($h, $p, $msg) { if ($h -notmatch $p) { throw "ASSERT FAILED: $msg (haystack='$h')" } }

InModuleScope -ModuleName QC.Processors {
    function Get-PWQcPrependRoleFieldsFromSourcePdf {
        param($FolderPath, $SourceDocumentName, $Config)
        return @{ found = $false; qcProcessType = ''; qcReviewType = '' }
    }
    function Get-PWQcPrependProcessIntentFromSourcePdf {
        param($FolderPath, $SourceDocumentName, $Config)
        if ($script:mockProcessIntent) { return $script:mockProcessIntent }
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
        if ($script:mockSheetIndexProcessType) {
            [void]$table.Rows.Add([string]$script:mockSheetIndexProcessType)
        }
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

    $initialIntake = @{
        id = 'j-initial-intake'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = '080J082001ab001.pdf'
        metadata = @{ prependTrigger = 'initialQcPdf'; triggerDocumentName = '080J082001ab001.pdf' }
    }
    $initialLane = _QCP-TryResolvePrependLaneContext -Job $initialIntake -Config $cfg
    Assert-True $initialLane.IsSuccess 'initialQcPdf on stem sheet defaults to production'
    Assert-Eq $initialLane.Data.qcProcessType 'production' 'initial intake => production'
    Assert-Eq $initialLane.Data.resolutionSource 'initial_prepend_default' 'initial intake resolution source'
    Assert-Eq $initialLane.Data.expectedLanePdfName '080J082001ab001-prod.pdf' 'initial intake creates *-prod.pdf'

    $finalPrependJob = @{
        id = 'j-final-rev-trigger'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = '080J082001ab001.pdf'
        metadata = @{
            triggerDocumentName = '080J082001ab001-rev.pdf'
            triggerDocumentGuid = '7362ac50-bf4c-4dfb-b4c5-4d4aac912ba4'
        }
    }
    $finalLane = _QCP-TryResolvePrependLaneContext -Job $finalPrependJob -Config $cfg
    Assert-True $finalLane.IsSuccess 'final prepend resolves lane from metadata trigger *-rev.pdf'
    Assert-Eq $finalLane.Data.qcProcessType 'review' 'lane trigger *-rev.pdf => review'
    Assert-Eq $finalLane.Data.expectedLanePdfName '080J082001ab001-rev.pdf' 'expected lane pdf from trigger suffix'
    Assert-Eq $finalLane.Data.resolutionSource 'document_name_lane' 'lane suffix wins over stem sourceName'

    $script:mockProcessIntent = @{
        found = $true
        qcProcessType = 'Review'
    }
    $reviewIntentJob = @{
        id = 'j-review-intent'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = '0818000063ea501.pdf'
        metadata = @{}
    }
    $reviewLane = _QCP-TryResolvePrependLaneContext -Job $reviewIntentJob -Config $cfg
    Assert-True $reviewLane.IsSuccess 'PDF QC_Process_Type review resolves over stale DGN production'
    Assert-Eq $reviewLane.Data.qcProcessType 'review' 'review lane from QC_Process_Type'
    Assert-Eq $reviewLane.Data.expectedLanePdfName '0818000063ea501-rev.pdf' 'expected *-rev.pdf from QC_Process_Type'
    $script:mockProcessIntent = $null

    $script:mockSheetIndexProcessType = 'production'
    $script:mockProcessIntent = @{
        found = $true
        qcProcessType = 'Check'
    }
    $staleIndexJob = @{
        id = 'j-stale-index-check'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = '045_D-01.01_d0847key.pdf'
        metadata = @{}
    }
    $checkLane = _QCP-TryResolvePrependLaneContext -Job $staleIndexJob -Config $cfg
    Assert-True $checkLane.IsSuccess 'live PW Check resolves over stale sheet_index production'
    Assert-Eq $checkLane.Data.qcProcessType 'check' 'check lane when PW disagrees with sheet_index'
    Assert-Eq $checkLane.Data.expectedLanePdfName '045_D-01.01_d0847key-chk.pdf' 'expected *-chk.pdf from live PW'
    Assert-Eq $checkLane.Data.resolutionSource 'projectwise_attributes' 'PW wins over stale sheet_index'
    $script:mockSheetIndexProcessType = $null
    $script:mockProcessIntent = $null

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

# scripts/processing/Invoke-QCPrependPw.ps1 honors lane parameters (no PW connect; exits after param resolution logging would need mock - use strict param block)
$pwPrependScript = Join-Path $repoRoot 'scripts\processing\Invoke-QCPrependPw.ps1'
Assert-True (Test-Path -LiteralPath $pwPrependScript) 'ProjectWise prepend script exists'
$pwPrependText = Get-Content -LiteralPath $pwPrependScript -Raw
Assert-Match $pwPrependText 'QcProcessType' 'prepend script defines QcProcessType'
Assert-Match $pwPrependText 'QcPdfSuffix' 'prepend script defines QcPdfSuffix'
Assert-Match $pwPrependText 'HistoryDocumentName' 'prepend script defines HistoryDocumentName'
Assert-Match $pwPrependText 'QC_PROCESS_TYPE_UNKNOWN' 'prepend script fails without lane in strict mode'

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
