# Phase 4: package-aware telemetry, reporting views, and correlation (mocked SQL).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

$phase4PackageId = [guid]::Parse('22222222-2222-2222-2222-222222222222')
$phase4DocGuid = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
$phase4Config = @{ database = @{ enabled = $true; connectionString = 'x' } }
$mockPackageId = $phase4PackageId
$mockDocGuid = $phase4DocGuid
$config = $phase4Config

# --- Schema v1.17 reporting views ---
$dbText = Get-Content -LiteralPath (Resolve-ModuleImplPath -ModuleName 'Core.Database.psm1') -Raw
Assert-True ($dbText -match '_QDB-GetSchemaV1dot17Additive') 'schema v1.17 additive patch exists'
Assert-True ($dbText -match 'CREATE VIEW v_sheet_package_status') 'v_sheet_package_status defined'
Assert-True ($dbText -match 'CREATE VIEW v_sheet_package_cycle_aging') 'v_sheet_package_cycle_aging defined'
Assert-True ($dbText -match "targetVersion = '1.19.0'") 'schema target version is 1.19.0'

InModuleScope -ModuleName Core.Database {
    function _QDB-IsEnabled { param([hashtable]$Config) return $true }
    function _QDB-NormalizeTelemetryPath { param([string]$Path) return ([string]$Path).Trim() }
    function Test-QCDatabaseWritesAllowed { param([hashtable]$Config) return $true }
    function Write-QCJsonLog { param([hashtable]$Data) }
    function Get-QCProcessingJobType { param([string]$QueueJobType, [hashtable]$Config = $null) return $QueueJobType }

    function Get-SheetPackageIdForDocument {
        param([hashtable]$Config, [string]$DocumentGuid)
        if ($DocumentGuid -eq $phase4DocGuid) { return $phase4PackageId }
        return $null
    }

    function Resolve-SheetPackageIdForSheetGroup {
        param([hashtable]$Config, [string]$FolderPath = '', [string]$SheetStem = '', [string]$DocumentGuid = '', [string]$DocumentName = '')
        return $phase4PackageId
    }

    function Resolve-SheetPackageFromDocument {
        param([string]$DocumentGuid = '', [string]$DocumentName = '', [string]$FolderPath = '')
        return @{
            documentGuid = $DocumentGuid
            documentName = $DocumentName
            folderPath = $FolderPath
            sheetStem = '080J082001ab001'
            documentRole = 'sheet_pdf'
            isSheetPackageMember = $true
        }
    }

    $script:lastJobSql = ''
    $script:lastJobParams = @{}
    function Invoke-QCDatabaseNonQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        $script:lastJobSql = $Sql
        $script:lastJobParams = $Parameters
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ rowsAffected = 1 }
    }

    # Processing jobs: sheet_package_id populated for QC_PREPEND
    $jobRes = Write-QCJobTelemetry -Config $phase4Config -JobId 'qc_test_job' -JobType 'QC_PREPEND' -Status 'succeeded' `
        -SourceFolder 'Documents\X\CADD\Sheets' -SourcePath 'Documents\X\CADD\Sheets\080J082001ab001.pdf' `
        -DocumentGuid $phase4DocGuid
    Assert-True $jobRes.IsSuccess 'Write-QCJobTelemetry should succeed'
    Assert-True ($script:lastJobSql -match 'sheet_package_id') 'processing_jobs MERGE includes sheet_package_id'
    Assert-Eq $script:lastJobParams.sheetPackageId $phase4PackageId 'processing_jobs.sheet_package_id populated'

    # Notifications: sheet_package_id populated from document_guid
    $script:lastNotifSql = ''
    $script:lastNotifParams = @{}
    function Invoke-QCDatabaseNonQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        if ($Sql -match 'notification_log') {
            $script:lastNotifSql = $Sql
            $script:lastNotifParams = $Parameters
        }
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ rowsAffected = 1 }
    }
    $notifRes = Write-QCNotificationTelemetry -Config $phase4Config -EventType 'READY_FOR_QC' `
        -DocumentGuid $phase4DocGuid -DocumentName '080J082001ab001.pdf' -FolderPath 'Documents\X\CADD\Sheets' `
        -Recipients 'a@example.com' -Subject 'Ready' -Success $true
    Assert-True $notifRes.IsSuccess 'Write-QCNotificationTelemetry should succeed'
    Assert-True ($script:lastNotifSql -match 'sheet_package_id') 'notification_log insert includes sheet_package_id'
    Assert-Eq $script:lastNotifParams.sheetPackageId $phase4PackageId 'notification_log.sheet_package_id populated'
}

# Reporting views: one row per package (static SQL shape)
$statusView = ($dbText -split '_QDB-GetSchemaV1dot17Additive', 2)[1] -split 'function _QDB-GetSchemaV1dot9Additive', 2 | Select-Object -First 1
Assert-True ($statusView -match 'FROM sheet_packages sp') 'v_sheet_package_status is package-grain'
Assert-True ($statusView -notmatch 'JOIN sheet_documents') 'v_sheet_package_status does not fan out to documents'
Assert-True ($statusView -match 'production_qc_completed_count') 'status view exposes completion counts'

$agingView = $statusView
Assert-True ($agingView -match 'v_sheet_package_cycle_aging') 'aging view present in v1.17 patch'
Assert-True ($agingView -match 'days_in_current_state') 'aging view exposes days_in_current_state'
Assert-True ($agingView -match 'days_since_last_completion') 'aging view exposes days_since_last_completion'

# QC.Reporting package metrics
Import-Module (Join-Path $repoRoot 'modules\Reporting\QC.Reporting.psm1') -Force
$table = New-Object System.Data.DataTable
foreach ($col in @('sheet_package_id','sheet_stem','folder_path','pw_state_name','qc_review_type','qc_assigned_to',
    'production_qc_completed_count','peer_review_completed_count','independent_check_completed_count',
    'production_qc_last_completed_at','peer_review_last_completed_at','independent_check_last_completed_at',
    'dgn_guid','sheet_pdf_guid','qc_pdf_guid')) {
    [void]$table.Columns.Add($col)
}
$row = $table.NewRow()
$row.sheet_package_id = $mockPackageId
$row.sheet_stem = '080J082001ab001'
$row.folder_path = 'Documents\X\CADD\Sheets'
$row.pw_state_name = 'Ready for QC'
$row.qc_review_type = 'Production QC'
$row.production_qc_completed_count = 1
$row.peer_review_completed_count = 0
$row.independent_check_completed_count = 0
$table.Rows.Add($row) | Out-Null

$dummyDoc = [pscustomobject]@{ Name = '080J082001ab001.pdf'; StateName = 'Ready for QC' }
$snap = New-QCReportingSnapshot -Documents @($dummyDoc) -Settings (Get-QCReportingSettings -Config $config) -PackageRows @($row)
Assert-Eq $snap.primaryEntity 'sheet_package_id' 'reporting primary entity is sheet_package_id'
Assert-Eq @($snap.packages).Count 1 'one row per package in reporting snapshot'
Assert-Eq $snap.packages[0].sheetStem '080J082001ab001' 'package record carries sheet stem'
Assert-Eq $snap.packageMetrics.packageCount 1 'package metrics count packages not documents'

Write-Host 'test_sheet_package_phase4.ps1 passed' -ForegroundColor Green
