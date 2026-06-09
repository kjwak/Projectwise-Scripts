# QC PDF GUID resolution prefers package member tables over ambiguous sheet_index rows.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/QC.Notifications.psm1') -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$liveGuid = '6f353f6f-ad61-4c01-96fe-56bd183b71a5'
$staleGuid = 'dfd34be2-56fd-486e-8ac3-d46967ac0df2'
$folder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$qcName = '0818000063ea500-qc.pdf'
$config = @{ database = @{ enabled = $true; connectionString = 'x' } }

InModuleScope -ModuleName QC.Notifications {
    function Test-QCDatabaseEnabled { param([hashtable]$Config) return $true }

    function Invoke-QCDatabaseQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})

        if ($Sql -match 'sheet_documents') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('document_guid', [string])
            [void]$table.Rows.Add($liveGuid)
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
        -QcPdfName $qcName -SheetStem '0818000063ea500' -HintGuid $staleGuid
    Assert-Eq $resolved.documentGuid $liveGuid 'package member GUID should win over stale hint and sheet_index'
    Assert-Eq $resolved.resolutionSource 'sheet_documents' 'resolution source should identify sheet_documents'

    $url = 'https://example.test/pwlink?objectId=6f353f6f-ad61-4c01-96fe-56bd183b71a5&objectType=doc&datasource=x&app=pwe'
    Assert-Eq (_QCN-ExtractPwLinkDocumentGuid -Url $url) $liveGuid 'pwlink objectId should be extracted from URL'
}

Write-Host 'OK: QC notification GUID resolution tests passed.' -ForegroundColor Green
