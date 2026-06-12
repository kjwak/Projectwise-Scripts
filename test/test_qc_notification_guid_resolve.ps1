# QC PDF GUID resolution prefers authoritative sheet_packages over stale sheet_documents.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/QC.Notifications.psm1') -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$liveGuid = '0f9c6ba8-a5e1-40ed-b2a3-2b906cb4f38b'
$staleGuid = '35b253e0-a6b7-44ec-a248-09e97d454d58'
$folder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$qcName = '080J082001ab001-qc.pdf'
$config = @{ database = @{ enabled = $true; connectionString = 'x' } }

InModuleScope -ModuleName QC.Notifications {
    function Test-QCDatabaseEnabled { param([hashtable]$Config) return $true }

    function Get-PWDocumentsBySearch {
        param([string]$FolderPath, [string]$DocumentName, [switch]$JustThisFolder)
        return @([pscustomobject]@{ DocumentGUID = $staleGuid; Name = $DocumentName })
    }
    function Get-PWDocumentsByGUIDs {
        param([string[]]$DocumentGUIDs)
        $g = [string]$DocumentGUIDs[0]
        if ($g -eq $liveGuid) {
            return @([pscustomobject]@{ Name = $qcName; DocumentGUID = $liveGuid })
        }
        return @()
    }

    function Invoke-QCDatabaseQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})

        if ($Sql -match 'sheet_package_qc_pdfs') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('document_guid', [string])
            [void]$table.Rows.Add($liveGuid)
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        if ($Sql -match 'FROM sheet_packages' -and $Sql -match 'qc_pdf_guid') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('lane_guid', [string])
            [void]$table.Rows.Add($liveGuid)
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        if ($Sql -match 'sheet_documents') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('document_guid', [string])
            [void]$table.Rows.Add($staleGuid)
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        if ($Sql -match 'FROM sheet_index' -and $Sql -match 'document_guid') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('document_guid', [string])
            [void]$table.Rows.Add($staleGuid)
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
    }

    $resolved = _QCN-ResolveLiveQcPdfDocumentGuidResult -Config $config -FolderPath $folder `
        -QcPdfName '080J082001ab001-prod.pdf' -SheetStem '080J082001ab001' -HintGuid $staleGuid `
        -Event @{ qcProcessType = 'production' }
    Assert-Eq $resolved.documentGuid $liveGuid 'lane QC PDF GUID should win over stale sheet_documents, PW search, hint, and sheet_index'
    Assert-Eq $resolved.resolutionSource 'sheet_package_qc_pdfs' 'resolution source should identify lane table'

    $url = 'https://example.test/pwlink?objectId=0f9c6ba8-a5e1-40ed-b2a3-2b906cb4f38b&objectType=doc&datasource=x&app=pwe'
    Assert-Eq (_QCN-ExtractPwLinkDocumentGuid -Url $url) $liveGuid 'pwlink objectId should be extracted from URL'
}

Write-Host 'OK: QC notification GUID resolution tests passed.' -ForegroundColor Green
