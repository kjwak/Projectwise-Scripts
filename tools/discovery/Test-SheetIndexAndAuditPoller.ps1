<#
.SYNOPSIS
Validates schema v1.1.0 migration (sheet_index table) and audit poller integration.

.DESCRIPTION
1. Initializes schema (creates sheet_index table if missing)
2. Tests Write-QCSheetIndex upsert
3. Tests Update-QCSheetQcPdf linking
4. Runs audit trail scan via Invoke-AuditTrailScan
5. Verifies sheet index population from audit events
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [int]$Hours = 1
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

$modulesDir = Join-Path $repoRoot 'modules'
foreach ($mod in @('Core.Results.psm1', 'Core.Runtime.psm1', 'Core.Database.psm1', 'PW.Connection.psm1', 'PW.AuditPoller.psm1')) {
    $modPath = Join-Path $modulesDir $mod
    if (-not (Test-Path -LiteralPath $modPath)) {
        Write-Host "ERROR: Module not found: $modPath" -ForegroundColor Red
        return
    }
    try {
        Import-Module $modPath -Force -ErrorAction Stop
        Write-Host "  Imported $mod OK" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to import $mod" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.InnerException) {
            Write-Host "  Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        return
    }
}
Write-Host "  Checking Test-QCDatabaseEnabled available: $(if (Get-Command Test-QCDatabaseEnabled -ErrorAction SilentlyContinue) { 'YES' } else { 'NO' })" -ForegroundColor Cyan

function _DeepHashtable ($obj) {
    if ($obj -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = _DeepHashtable $p.Value }
        return $h
    }
    if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
        return @($obj | ForEach-Object { _DeepHashtable $_ })
    }
    return $obj
}

Write-Host "Loading config from: $AppSettingsPath" -ForegroundColor Cyan
$raw = Get-Content -LiteralPath $AppSettingsPath -Raw -ErrorAction Stop
$config = _DeepHashtable ($raw | ConvertFrom-Json -ErrorAction Stop)

$dbEnabled = Test-QCDatabaseEnabled -Config $config
Write-Host "  database.enabled = $dbEnabled" -ForegroundColor $(if ($dbEnabled) { 'Green' } else { 'Red' })
if (-not $dbEnabled) { Write-Host "Database not enabled. Exiting." -ForegroundColor Red; return }

# --- 1. Schema initialization ---
Write-Host "`n[1] Initializing schema (v1.1.0 with sheet_index)..." -ForegroundColor Yellow
$schemaRes = Initialize-QCDatabaseSchema -Config $config
Write-Host "  $($schemaRes.Code): $($schemaRes.Message)" -ForegroundColor $(if ($schemaRes.IsSuccess) { 'Green' } else { 'Red' })

# Verify sheet_index table exists
$tableCheck = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT CASE WHEN OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL THEN 1 ELSE 0 END"
$tableExists = $tableCheck.IsSuccess -and [int]$tableCheck.Data.value -eq 1
Write-Host "  sheet_index table exists: $tableExists" -ForegroundColor $(if ($tableExists) { 'Green' } else { 'Red' })

$viewCheck = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT CASE WHEN OBJECT_ID('dbo.v_sheet_status', 'V') IS NOT NULL THEN 1 ELSE 0 END"
$viewExists = $viewCheck.IsSuccess -and [int]$viewCheck.Data.value -eq 1
Write-Host "  v_sheet_status view exists: $viewExists" -ForegroundColor $(if ($viewExists) { 'Green' } else { 'Red' })

# --- 2. Test Write-QCSheetIndex ---
Write-Host "`n[2] Testing Write-QCSheetIndex (upsert)..." -ForegroundColor Yellow
$testGuid = 'test-' + [guid]::NewGuid().ToString('N').Substring(0, 20)
Write-QCSheetIndex -Config $config -DocumentGuid $testGuid `
    -DocumentName 'TEST_SHEET_001.pdf' -FolderPath 'AZDOT 2024\TestProject\CADD\Sheets' `
    -SourceType 'pdf' -DesignerEmail 'designer@test.com' -ReviewerEmail 'reviewer@test.com' `
    -PwStateName 'In Production' -WatchRoot 'Documents\AZDOT 2024'

$insertCheck = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT * FROM sheet_index WHERE document_guid = @g" -Parameters @{ g = $testGuid }
if ($insertCheck.IsSuccess -and $insertCheck.Data.rowCount -gt 0) {
    Write-Host "  INSERT OK: document_name=$([string]$insertCheck.Data.table.Rows[0].document_name), designer=$([string]$insertCheck.Data.table.Rows[0].designer_email)" -ForegroundColor Green
} else {
    Write-Host "  INSERT FAILED" -ForegroundColor Red
}

# Update the same row
Write-QCSheetIndex -Config $config -DocumentGuid $testGuid `
    -DocumentName 'TEST_SHEET_001.pdf' -FolderPath 'AZDOT 2024\TestProject\CADD\Sheets' `
    -QcStage 'Red' -QcStatus 'Open'

$updateCheck = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT qc_stage, qc_status FROM sheet_index WHERE document_guid = @g" -Parameters @{ g = $testGuid }
if ($updateCheck.IsSuccess -and $updateCheck.Data.rowCount -gt 0) {
    $row = $updateCheck.Data.table.Rows[0]
    Write-Host "  UPDATE OK: qc_stage=$([string]$row.qc_stage), qc_status=$([string]$row.qc_status)" -ForegroundColor Green
} else {
    Write-Host "  UPDATE FAILED" -ForegroundColor Red
}

# --- 3. Test Update-QCSheetQcPdf ---
Write-Host "`n[3] Testing Update-QCSheetQcPdf..." -ForegroundColor Yellow
Update-QCSheetQcPdf -Config $config -SourceDocumentGuid $testGuid `
    -QcPdfGuid 'qc-pdf-test-guid' -QcPdfName 'TEST_SHEET_001-qc.pdf'

$qcCheck = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT qc_pdf_guid, qc_pdf_name FROM sheet_index WHERE document_guid = @g" -Parameters @{ g = $testGuid }
if ($qcCheck.IsSuccess -and $qcCheck.Data.rowCount -gt 0) {
    $row = $qcCheck.Data.table.Rows[0]
    Write-Host "  QC PDF link OK: qc_pdf_name=$([string]$row.qc_pdf_name)" -ForegroundColor Green
} else {
    Write-Host "  QC PDF link FAILED" -ForegroundColor Red
}

# Clean up test row
Invoke-QCDatabaseNonQuery -Config $config -Sql "DELETE FROM sheet_index WHERE document_guid = @g" -Parameters @{ g = $testGuid } | Out-Null
Write-Host "  Test row cleaned up." -ForegroundColor Gray

# --- 4. Test v_sheet_status view ---
Write-Host "`n[4] Querying v_sheet_status..." -ForegroundColor Yellow
$viewRes = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT TOP 5 document_name, folder_path, designer_email, pw_state_name, has_qc_pdf FROM v_sheet_status ORDER BY last_updated_at DESC"
if ($viewRes.IsSuccess) {
    Write-Host "  v_sheet_status rows: $($viewRes.Data.rowCount)" -ForegroundColor Green
    foreach ($row in @($viewRes.Data.table.Rows | Select-Object -First 5)) {
        Write-Host "    $([string]$row.document_name) | $([string]$row.folder_path) | qc=$([string]$row.has_qc_pdf)" -ForegroundColor Cyan
    }
} else {
    Write-Host "  View query failed: $($viewRes.Message)" -ForegroundColor Red
}

# --- 5. Audit trail scan (if PW is available) ---
Write-Host "`n[5] Testing Invoke-AuditTrailScan..." -ForegroundColor Yellow
try {
    $pw = $config.projectWise
    $ds = if ($pw.datasourceName) { [string]$pw.datasourceName } elseif ($pw.datasource) { [string]$pw.datasource } else { $null }
    $credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }

    $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
    if (-not $credRes.IsSuccess) { throw "Credential: $($credRes.Message)" }
    $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
    if (-not $connRes.IsSuccess) { throw "Connection: $($connRes.Message)" }
    Write-Host "  Connected to PW: $ds" -ForegroundColor Green

    $since = (Get-Date).AddHours(-$Hours)
    Write-Host "  Scanning since: $($since.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray

    $scanRes = Invoke-AuditTrailScan -Config $config -Since $since
    if ($scanRes.IsSuccess) {
        $st = $scanRes.Data.stats
        Write-Host "  Total events:     $($st.totalEvents)" -ForegroundColor Cyan
        Write-Host "  QC-relevant:      $($st.relevantEvents)" -ForegroundColor Cyan
        Write-Host "  Folders resolved:  $($st.foldersResolved)" -ForegroundColor Cyan
        Write-Host "  Watch matches:    $($st.watchMatches)" -ForegroundColor $(if ($st.watchMatches -gt 0) { 'Green' } else { 'Yellow' })
        Write-Host "  Sheets matches:   $($st.sheetsMatches)" -ForegroundColor $(if ($st.sheetsMatches -gt 0) { 'Green' } else { 'Yellow' })
        Write-Host "  DB writes:        $($st.dbWrites)" -ForegroundColor Cyan
        Write-Host "  Watermark after:  $($scanRes.Data.watermarkAfter)" -ForegroundColor Gray
        Write-Host "  Duration:         $($scanRes.Data.durationMs)ms" -ForegroundColor Gray
    } else {
        Write-Host "  Scan failed: $($scanRes.Message)" -ForegroundColor Red
    }

    # 5b. Populate sheet_index from sheets candidates
    $sheetsCandidates = @($scanRes.Data.candidates | Where-Object { $_.isSheetsFolder })
    if ($sheetsCandidates.Count -gt 0) {
        Write-Host "`n[5b] Populating sheet_index from $($sheetsCandidates.Count) sheets candidates..." -ForegroundColor Yellow
        $sheetGuids = @($sheetsCandidates | ForEach-Object { $_.objGuid } | Select-Object -Unique)
        $indexed = 0
        foreach ($chunk in @(for ($ci = 0; $ci -lt $sheetGuids.Count; $ci += 200) { ,@($sheetGuids[$ci..[Math]::Min($ci + 199, $sheetGuids.Count - 1)]) })) {
            $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs $chunk -ErrorAction SilentlyContinue)
            foreach ($doc in $docs) {
                $dg = [string]$doc.DocumentGUID
                $folder = if ($scanRes.Data.docToFolder.ContainsKey($dg)) { $scanRes.Data.docToFolder[$dg] } else { $null }
                if (-not $folder) { continue }
                $ext = [System.IO.Path]::GetExtension([string]$doc.FileName)
                $designerEmail = $null; $reviewerEmail = $null; $stateName = $null
                try {
                    $attrs = $doc | Get-PWDocumentAttributes -ErrorAction SilentlyContinue
                    if ($attrs) {
                        $de = $attrs | Where-Object { $_.Name -eq 'EM_Designer_Email' } | Select-Object -First 1
                        $re = $attrs | Where-Object { $_.Name -eq 'EM_Reviewer_Email' } | Select-Object -First 1
                        if ($de) { $designerEmail = [string]$de.Value }
                        if ($re) { $reviewerEmail = [string]$re.Value }
                    }
                } catch { }
                try { $stateName = [string]$doc.StateName } catch { }

                Write-QCSheetIndex -Config $config -DocumentGuid $dg `
                    -DocumentName ([string]$doc.FileName) -FolderPath $folder `
                    -Extension $ext -SourceType ($ext -replace '^\.', '') `
                    -DesignerEmail $designerEmail -ReviewerEmail $reviewerEmail `
                    -PwStateName $stateName -WatchRoot ''
                $indexed++
            }
        }
        Write-Host "  Indexed $indexed sheets into sheet_index." -ForegroundColor Green
    }

    Disconnect-PW | Out-Null
} catch {
    Write-Host "  PW not available: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  (Skipping audit scan test -- DB tests above still valid)" -ForegroundColor Gray
}

# --- 6. Final sheet_index stats ---
Write-Host "`n[6] Sheet index stats..." -ForegroundColor Yellow
$countRes = Invoke-QCDatabaseScalar -Config $config -Sql "SELECT COUNT(*) FROM sheet_index"
if ($countRes.IsSuccess) {
    Write-Host "  Total sheet_index rows: $($countRes.Data.value)" -ForegroundColor Green
}
$byExt = Invoke-QCDatabaseQuery -Config $config -Sql "SELECT extension, COUNT(*) AS cnt FROM sheet_index GROUP BY extension ORDER BY cnt DESC"
if ($byExt.IsSuccess -and $byExt.Data.rowCount -gt 0) {
    Write-Host "  By extension:" -ForegroundColor Cyan
    foreach ($row in $byExt.Data.table.Rows) {
        Write-Host "    $([string]$row.extension): $([int]$row.cnt)" -ForegroundColor Gray
    }
}

Write-Host "`n=== All tests complete ===" -ForegroundColor Green
