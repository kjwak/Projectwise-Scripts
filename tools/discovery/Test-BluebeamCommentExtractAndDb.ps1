param(
    # If the PDF is local, pass this. If the PDF is only in ProjectWise, omit this and we will export by GUID.
    [string]$LocalPdfPath = '',

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
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force -ErrorAction SilentlyContinue

function Assert-True($Condition, $Message) { if (-not $Condition) { throw "ASSERT FAILED: $Message" } }

$configRes = Read-AppConfig -Path $ConfigPath
Assert-True $configRes.IsSuccess "Config load failed: $($configRes.Message)"
$config = $configRes.Data.config

if (-not $config.database) { $config['database'] = @{} }
$config.database.enabled = $true

Write-Host "[1] Initializing DB schema (if needed)..." -ForegroundColor Yellow
$schemaRes = Initialize-QCDatabaseSchema -Config $config
Assert-True $schemaRes.IsSuccess "Schema init failed: $($schemaRes.Message)"
Write-Host ("  Schema: " + $schemaRes.Message) -ForegroundColor Green

$pdfPathToUse = $LocalPdfPath
if ([string]::IsNullOrWhiteSpace($pdfPathToUse)) {
    # Export the *-qc.pdf from ProjectWise to a temp folder, by GUID.
    Assert-True ([bool](Get-Command -Name 'Invoke-PWAuthenticatedCommand' -ErrorAction SilentlyContinue)) `
        'PW.Connection is not available (Invoke-PWAuthenticatedCommand missing). Run on a machine with ProjectWise PowerShell installed.'
    Assert-True ($config.projectWise -and $config.projectWise.datasourceName -and $config.projectWise.credentialPath) `
        'Config missing projectWise.datasourceName / projectWise.credentialPath (required to export from ProjectWise).'

    $tmp = Join-Path $env:TEMP ("qc-bluebeam-export-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null

    Write-Host "[2] Exporting QC PDF from ProjectWise by GUID..." -ForegroundColor Yellow
    $ds = [string]$config.projectWise.datasourceName
    $credPath = [string]$config.projectWise.credentialPath
    $script:qcGuid = [string]$QcPdfGuid
    $script:outDir = $tmp
    $exportRes = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -ScriptBlock {
        $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
        if (-not $guidCmd) { throw 'Get-PWDocumentsByGUIDs not available.' }
        $doc = @(& $guidCmd -DocumentGUIDs @($script:qcGuid) -ErrorAction Stop) | Select-Object -First 1
        if (-not $doc) { throw "QC PDF GUID not found: $($script:qcGuid)" }

        $expCmd = Get-Command -Name 'Export-PWDocumentsSimple' -ErrorAction SilentlyContinue
        if (-not $expCmd) { throw 'Export-PWDocumentsSimple not available (pwps_dab missing).' }
        Export-PWDocumentsSimple -InputDocuments $doc -TargetFolder $script:outDir -ErrorAction Stop | Out-Null

        $name = $null
        try { $name = [string]$doc.Name } catch { }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = '*.pdf' }

        $f = Get-ChildItem -LiteralPath $script:outDir -File -Filter $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $f) {
            $f = Get-ChildItem -LiteralPath $script:outDir -File -Filter '*.pdf' -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
        }
        if (-not $f) { throw "Export succeeded but no PDF appeared in: $($script:outDir)" }
        return $f.FullName
    }
    $pdfPathToUse = [string]$exportRes
    Assert-True (-not [string]::IsNullOrWhiteSpace($pdfPathToUse)) 'ProjectWise export did not return a local path.'
    Assert-True (Test-Path -LiteralPath $pdfPathToUse) "Exported PDF not found: $pdfPathToUse"
    Write-Host ("  Exported to: " + $pdfPathToUse) -ForegroundColor Green
}

if (-not [string]::IsNullOrWhiteSpace($SheetGuid)) {
    if ([string]::IsNullOrWhiteSpace($SheetName)) { $SheetName = 'ManualSheet.pdf' }
    if ([string]::IsNullOrWhiteSpace($SheetFolderPath)) { $SheetFolderPath = 'Documents\\Manual\\Sheets' }

    Write-Host "[3] Upserting sheet_index + linking qc_pdf_guid..." -ForegroundColor Yellow
    Write-QCSheetIndex -Config $config -DocumentGuid $SheetGuid -DocumentName $SheetName -FolderPath $SheetFolderPath -SourceType 'PW'
    Update-QCSheetQcPdf -Config $config -SourceDocumentGuid $SheetGuid -QcPdfGuid $QcPdfGuid -QcPdfName ([System.IO.Path]::GetFileName($pdfPathToUse))
    Write-Host "  sheet_index linked." -ForegroundColor Green
}

Write-Host "[4] Extracting comments (Python parser)..." -ForegroundColor Yellow
$extractRes = Invoke-QCCommentExtract -LocalPdfPath $pdfPathToUse -Config $config
Assert-True $extractRes.IsSuccess "Extract failed: $($extractRes.Message)"
$ann = @($extractRes.Data.annotations)
Write-Host ("  Extracted: {0} annotation(s), parser_status={1}, parser_version={2}" -f $ann.Count, $extractRes.Data.parserStatus, $extractRes.Data.parserVersion) -ForegroundColor Green
Assert-True ($ann.Count -gt 0) "No annotations found. Confirm the PDF actually contains Bluebeam markups/comments."

Write-Host "[5] Writing a qc_comment_runs row + qc_comments snapshots..." -ForegroundColor Yellow
$jobId = 'manual-' + [guid]::NewGuid().ToString('N')
$runRecord = @{
    job_id = $jobId
    document_id = $QcPdfGuid
    project_id = $null
    pw_path = $null
    file_name = [System.IO.Path]::GetFileName($pdfPathToUse)
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

Write-Host "[6] Verifying rows in DB..." -ForegroundColor Yellow
$q = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT COUNT(*) FROM qc_comments WHERE run_id = @r" -Parameters @{ r = $runId }
Assert-True $q.IsSuccess "DB query failed: $($q.Message)"
Assert-True ([int]$q.Data.value -gt 0) 'Expected qc_comments rows for this run'
Write-Host ("  qc_comments rows for run_id {0}: {1}" -f $runId, $q.Data.value) -ForegroundColor Green

Write-Host "OK Test-BluebeamCommentExtractAndDb.ps1" -ForegroundColor Green

