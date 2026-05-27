param(
    [Parameter(Mandatory)]
    [string]$LocalPdfPath,

    # The ProjectWise GUID for the *-qc.pdf document (used as qc_comments.document_id)
    [Parameter(Mandatory)]
    [string]$QcPdfGuid,

    # Optional: the source sheet GUID; if provided we will upsert sheet_index + link qc_pdf_guid.
    [string]$SheetGuid = '',
    [string]$SheetName = '',
    [string]$SheetFolderPath = '',

    # Optional: override config path (defaults to repo appsettings.json)
    [string]$ConfigPath = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Config.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.CommentExtract.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.CommentSync.Database.psm1') -Force

function Assert-True($Condition, $Message) { if (-not $Condition) { throw "ASSERT FAILED: $Message" } }

if (-not (Test-Path -LiteralPath $LocalPdfPath)) {
    throw "PDF not found: $LocalPdfPath"
}

$configRes = Read-AppConfig -Path $ConfigPath
Assert-True $configRes.IsSuccess "Config load failed: $($configRes.Message)"
$config = $configRes.Data.config

if (-not $config.database) { $config['database'] = @{} }
$config.database.enabled = $true

Write-Host "[1] Initializing DB schema (if needed)..." -ForegroundColor Yellow
$schemaRes = Initialize-QCDatabaseSchema -Config $config
Assert-True $schemaRes.IsSuccess "Schema init failed: $($schemaRes.Message)"
Write-Host ("  Schema: " + $schemaRes.Message) -ForegroundColor Green

if (-not [string]::IsNullOrWhiteSpace($SheetGuid)) {
    if ([string]::IsNullOrWhiteSpace($SheetName)) { $SheetName = 'ManualSheet.pdf' }
    if ([string]::IsNullOrWhiteSpace($SheetFolderPath)) { $SheetFolderPath = 'Documents\\Manual\\Sheets' }

    Write-Host "[2] Upserting sheet_index + linking qc_pdf_guid..." -ForegroundColor Yellow
    Write-QCSheetIndex -Config $config -DocumentGuid $SheetGuid -DocumentName $SheetName -FolderPath $SheetFolderPath -SourceType 'PW'
    Update-QCSheetQcPdf -Config $config -SourceDocumentGuid $SheetGuid -QcPdfGuid $QcPdfGuid -QcPdfName ([System.IO.Path]::GetFileName($LocalPdfPath))
    Write-Host "  sheet_index linked." -ForegroundColor Green
}

Write-Host "[3] Extracting comments (Python parser)..." -ForegroundColor Yellow
$extractRes = Invoke-QCCommentExtract -LocalPdfPath $LocalPdfPath -Config $config
Assert-True $extractRes.IsSuccess "Extract failed: $($extractRes.Message)"
$ann = @($extractRes.Data.annotations)
Write-Host ("  Extracted: {0} annotation(s), parser_status={1}, parser_version={2}" -f $ann.Count, $extractRes.Data.parserStatus, $extractRes.Data.parserVersion) -ForegroundColor Green
Assert-True ($ann.Count -gt 0) "No annotations found. Confirm the PDF actually contains Bluebeam markups/comments."

Write-Host "[4] Writing a qc_comment_runs row + qc_comments snapshots..." -ForegroundColor Yellow
$jobId = 'manual-' + [guid]::NewGuid().ToString('N')
$runRecord = @{
    job_id = $jobId
    document_id = $QcPdfGuid
    project_id = $null
    pw_path = $null
    file_name = [System.IO.Path]::GetFileName($LocalPdfPath)
    file_hash = $null
    source_modified_utc = $null
    previous_pw_state = $null
    target_pw_state = $null
    state_update_result = $null
    parser_status = [string]$extractRes.Data.parserStatus
    processor_version = 'manual-test'
}

$runRes = Write-QCCommentSyncRun -Config $config -RunRecord $runRecord
Assert-True $runRes.IsSuccess "Run insert failed: $($runRes.Message)"
$runId = [long]$runRes.Data.runId
Assert-True ($runId -gt 0) 'Expected run_id > 0'

$writeComments = Write-QCCommentRows -Config $config -RunId $runId -DocumentId $QcPdfGuid -Annotations $ann
Assert-True $writeComments.IsSuccess "Comments insert failed: $($writeComments.Message)"
Write-Host ("  " + $writeComments.Message) -ForegroundColor Green

$writeHist = Write-QCCommentStatusHistoryRows -Config $config -RunId $runId -DocumentId $QcPdfGuid -Annotations $ann
Assert-True $writeHist.IsSuccess "History insert failed: $($writeHist.Message)"
Write-Host ("  " + $writeHist.Message) -ForegroundColor Green

Write-Host "[5] Verifying rows in DB..." -ForegroundColor Yellow
$q = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT COUNT(*) FROM qc_comments WHERE run_id = @r" -Parameters @{ r = $runId }
Assert-True $q.IsSuccess "DB query failed: $($q.Message)"
Assert-True ([int]$q.Data.value -gt 0) 'Expected qc_comments rows for this run'
Write-Host ("  qc_comments rows for run_id {0}: {1}" -f $runId, $q.Data.value) -ForegroundColor Green

Write-Host "OK Test-BluebeamCommentExtractAndDb.ps1" -ForegroundColor Green

