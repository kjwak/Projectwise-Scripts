# Lane QC PDF GUID: live PW search wins over stale DB telemetry.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules/Core.Database.psm1') -Force

function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }

$liveGuid = 'c93781c5-4540-4916-a97e-5a07a7f70433'
$staleGuid = '00913b00-94c6-4c23-b6a3-7b379b527141'
$folder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$apiFolder = 'caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$chkName = '080J082001ca001-chk.pdf'
$config = @{ database = @{ enabled = $true; connectionString = 'x' } }

InModuleScope -ModuleName Core.Database {
    function Test-QCDatabaseEnabled { param([hashtable]$Config) return $true }
    function _QDB-IsEnabled { param([hashtable]$Config) return $true }
    function _QDB-NormalizeTelemetryPath { param([string]$Path) return $Path }

    function Get-PWDocumentsBySearch {
        param([string]$FolderPath, [string]$DocumentName, [switch]$JustThisFolder)
        if ($FolderPath -eq $apiFolder -and $DocumentName -eq $chkName) {
            return @([pscustomobject]@{ DocumentGUID = $liveGuid; Name = $DocumentName; FileUpdatedDate = (Get-Date) })
        }
        return @()
    }

    function Invoke-QCDatabaseQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        if ($Sql -match 'sheet_package_qc_pdfs') {
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
        if ($Sql -match 'qc_chk_pdf_guid') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('lane_guid', [string])
            [void]$table.Rows.Add($staleGuid)
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
    }

    $resolved = Resolve-QCSheetQcPdfGuid -Config $config -FolderPath $folder -QcPdfName $chkName
    Assert-Eq $resolved $liveGuid 'Resolve-QCSheetQcPdfGuid should prefer live PW search over stale DB'
}

Write-Host 'OK: lane PDF GUID sync tests passed.' -ForegroundColor Green
