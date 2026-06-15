# Tests sheet-level QC_PREPEND dedupe (blocks duplicate enqueue while job is active).
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/QC.ProcessType.psm1" -Force
Import-Module "$repoRoot/modules/QC.Processors.psm1" -Force

function New-TempQueueRoot {
    $t = Join-Path $env:TEMP ("qc-prepend-dedupe-" + ([guid]::NewGuid().ToString('N')))
    foreach ($sub in @('pending', 'running', 'succeeded', 'failed')) {
        New-Item -ItemType Directory -Path (Join-Path $t $sub) -Force | Out-Null
    }
    return $t
}

function Write-PrependQueueJob {
    param(
        [string]$QueueRoot,
        [string]$State,
        [string]$JobId,
        [string]$FolderPath,
        [string]$SheetPdfName,
        [string]$DedupeKey = 'dq_test',
        [string]$QcProcessType = '',
        [string]$UpdatedAtUtc = ''
    )
    if (-not $UpdatedAtUtc) { $UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o') }
    $metadata = @{
        prependTrigger = 'initialQcPdf'
        stateTransitionKey = 'audit:999'
    }
    if (-not [string]::IsNullOrWhiteSpace($QcProcessType)) { $metadata['qcProcessType'] = $QcProcessType }
    $job = @{
        id = $JobId
        type = 'QC_PREPEND'
        state = $State
        dedupeKey = $DedupeKey
        sourceFolder = $FolderPath
        sourceName = $SheetPdfName
        sourcePath = (Join-Path $FolderPath $SheetPdfName)
        metadata = $metadata
        updatedAtUtc = $UpdatedAtUtc
    }
    $path = Join-Path (Join-Path $QueueRoot $State) ($JobId + '.json')
    $job | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $path -Encoding UTF8
}

$queueRoot = New-TempQueueRoot
$folder = 'C:\pw\test\Sheets\080J082001'
$sheetPdf = '080J082001ca001.pdf'
$config = @{ queue = @{ rootDir = $queueRoot } }

try {
    $open = Test-QCPrependEnqueueBlockedForSheet -Config $config -FolderPath $folder -SheetPdfName $sheetPdf
    Assert-True (-not $open.blocked) 'No jobs: should not block'

    Write-PrependQueueJob -QueueRoot $queueRoot -State 'running' -JobId 'qc_qcprepend_running1' `
        -FolderPath $folder -SheetPdfName $sheetPdf
    $running = Test-QCPrependEnqueueBlockedForSheet -Config $config -FolderPath $folder -SheetPdfName $sheetPdf
    Assert-True $running.blocked 'Running job should block'
    Assert-Eq $running.reason 'queue_running' 'Running reason'

    Remove-Item -LiteralPath (Join-Path $queueRoot 'running\qc_qcprepend_running1.json') -Force
    $recent = (Get-Date).ToUniversalTime().AddMinutes(-2).ToString('o')
    Write-PrependQueueJob -QueueRoot $queueRoot -State 'succeeded' -JobId 'qc_qcprepend_done1' `
        -FolderPath $folder -SheetPdfName $sheetPdf -UpdatedAtUtc $recent
    $succeeded = Test-QCPrependEnqueueBlockedForSheet -Config $config -FolderPath $folder -SheetPdfName $sheetPdf
    Assert-True $succeeded.blocked 'Recent succeeded job should block'
    Assert-Eq $succeeded.reason 'queue_succeeded_recent' 'Recent succeeded reason'

    $old = (Get-Date).ToUniversalTime().AddHours(-2).ToString('o')
    Remove-Item -LiteralPath (Join-Path $queueRoot 'succeeded\qc_qcprepend_done1.json') -Force
    Write-PrependQueueJob -QueueRoot $queueRoot -State 'succeeded' -JobId 'qc_qcprepend_old1' `
        -FolderPath $folder -SheetPdfName $sheetPdf -UpdatedAtUtc $old
    $stale = Test-QCPrependEnqueueBlockedForSheet -Config $config -FolderPath $folder -SheetPdfName $sheetPdf
    Assert-True (-not $stale.blocked) 'Old succeeded job should not block'

    Remove-Item -LiteralPath (Join-Path $queueRoot 'succeeded\qc_qcprepend_old1.json') -Force
    $recentProd = (Get-Date).ToUniversalTime().AddMinutes(-2).ToString('o')
    Write-PrependQueueJob -QueueRoot $queueRoot -State 'succeeded' -JobId 'qc_qcprepend_prod1' `
        -FolderPath $folder -SheetPdfName $sheetPdf -QcProcessType 'production' -UpdatedAtUtc $recentProd
    $prodBlocked = Test-QCPrependEnqueueBlockedForSheet -Config $config -FolderPath $folder -SheetPdfName $sheetPdf `
        -QcProcessType 'production'
    Assert-True $prodBlocked.blocked 'Recent production succeeded job should block production lane'
    $checkOpen = Test-QCPrependEnqueueBlockedForSheet -Config $config -FolderPath $folder -SheetPdfName $sheetPdf `
        -QcProcessType 'check'
    Assert-True (-not $checkOpen.blocked) 'Check lane should not be blocked by recent production succeeded job'

    Write-Host 'test_qc_prepend_sheet_dedupe.ps1: all assertions passed.'
}
finally {
    Remove-Item -LiteralPath $queueRoot -Recurse -Force -ErrorAction SilentlyContinue
}
