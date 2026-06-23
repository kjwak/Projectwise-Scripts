# Lane trigger resolution: source PDF + QC_Process_Type and direct lane PDF triggers.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.Processors.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.Notifications.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }

InModuleScope -ModuleName QC.Processors {
    function Get-PWQcPrependRoleFieldsFromSourcePdf {
        param($FolderPath, $SourceDocumentName, $Config)
        return @{
            found = $true
            qcProcessType = 'check'
            qcReviewType = 'Independent Check'
            designerEmail = 'a@example.com'
            reviewerEmail = ''
            checkerEmail = ''
        }
    }
    function Get-PWQcPdfLaneFromDocumentName {
        param([string]$DocumentName)
        if ($DocumentName -match '(?i)-chk\.pdf$') { return 'check' }
        if ($DocumentName -match '(?i)-rev\.pdf$') { return 'review' }
        if ($DocumentName -match '(?i)-prod\.pdf$') { return 'production' }
        return $null
    }
    function Get-PWQcDefaultProcessType { param($Config) return 'production' }
    function Invoke-QCDatabaseQuery {
        param($Config, $Sql, $Parameters = @{})
        $table = New-Object System.Data.DataTable
        [void]$table.Columns.Add('qc_process_type', [string])
        [void]$table.Rows.Add('check')
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ table = $table }
    }
    function Test-QCDatabaseEnabled { param($Config) return $true }

    $jobSource = @{
        id = 'j1'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = 'CA001.pdf'
        metadata = @{ qcProcessType = 'check' }
    }
    Assert-Eq (_QCP-ResolveProcessTypeFromJob -Job $jobSource -Config @{}) 'check' 'source PDF + metadata qcProcessType resolves check'

    $jobDirectChk = @{
        id = 'j2'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = 'CA001-chk.pdf'
    }
    Assert-Eq (_QCP-ResolveProcessTypeFromJob -Job $jobDirectChk -Config @{}) 'check' 'direct *-chk.pdf trigger resolves check'

    $jobDirectRev = @{
        id = 'j3'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = 'CA001-rev.pdf'
    }
    Assert-Eq (_QCP-ResolveProcessTypeFromJob -Job $jobDirectRev -Config @{}) 'review' 'direct *-rev.pdf trigger resolves review'

    $jobSheetIndex = @{
        id = 'j4'
        sourceFolder = 'Documents\X\CADD\Sheets'
        sourceName = 'CA001.pdf'
    }
    Assert-Eq (_QCP-ResolveProcessTypeFromJob -Job $jobSheetIndex -Config @{ database = @{ enabled = $true } }) 'check' 'stem PDF resolves qc_process_type from sheet_index'
}

InModuleScope -ModuleName QC.Notifications {
    function Test-QCDatabaseEnabled { param($Config) return $true }
    function Get-PWDocumentsBySearch { param($FolderPath, $DocumentName) return @() }
    function Get-PWDocumentsByGUIDs { param($DocumentGUIDs) return @() }
    function Invoke-QCDatabaseQuery {
        param($Config, $Sql, $Parameters = @{})
        $table = New-Object System.Data.DataTable
        if ($Sql -match 'sheet_package_qc_pdfs' -and $Parameters.qcProcessType -eq 'check') {
            [void]$table.Columns.Add('document_guid', [string])
            [void]$table.Rows.Add('22222222-2222-2222-2222-222222222222')
            return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ table = $table }
        }
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ table = $null }
    }

    $event = @{
        qcProcessType = 'check'
        documentName = 'CA001.pdf'
        folderPath = 'Documents\X\CADD\Sheets'
    }
    $resolved = _QCN-ResolveLiveQcPdfDocumentGuidResult -Config @{ database = @{ enabled = $true } } `
        -FolderPath 'Documents\X\CADD\Sheets' -QcPdfName 'CA001-chk.pdf' -SheetStem 'CA001' -Event $event
    Assert-Eq $resolved.documentGuid '22222222-2222-2222-2222-222222222222' 'notification link resolves active check lane GUID'
    Assert-Eq $resolved.resolutionSource 'sheet_package_qc_pdfs' 'notification uses lane table not production alias'
}

Write-Host 'test_qc_lane_trigger_resolution: OK' -ForegroundColor Green
