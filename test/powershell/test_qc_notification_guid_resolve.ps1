# QC PDF GUID resolution prefers authoritative sheet_packages over stale sheet_documents.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/ProjectWise/PW.Discovery.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Notifications/QC.Notifications.psm1') -Force

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

$liveGuid = '0f9c6ba8-a5e1-40ed-b2a3-2b906cb4f38b'
$staleGuid = '35b253e0-a6b7-44ec-a248-09e97d454d58'
$chkGuid = 'fcbe7d2c-1111-2222-3333-444455556666'
$folder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$apiFolder = 'caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
$qcName = '080J082001ab001-qc.pdf'
$chkName = '0818000063ea501-chk.pdf'
$lowerChkName = '080j082001ab001-chk.pdf'
$mixedChkGuid = 'bc4b38fa-6333-442d-9b63-e63bfd2cafe6'
$config = @{ database = @{ enabled = $true; connectionString = 'x' } }

InModuleScope -ModuleName QC.Notifications {
    function Test-QCDatabaseEnabled { param([hashtable]$Config) return $true }

    $script:pwSearchFolders = @()
    function Get-PWDocumentsBySearch {
        param([string]$FolderPath, [string]$DocumentName, [switch]$JustThisFolder)
        $script:pwSearchFolders += $FolderPath
        if ($FolderPath -eq $apiFolder) {
            if ($DocumentName -ieq $lowerChkName) { return @() }
            if ($DocumentName -eq $chkName) {
                return @([pscustomobject]@{ DocumentGUID = $chkGuid; Name = $DocumentName })
            }
        }
        if ($FolderPath -eq $folder) { return @() }
        return @([pscustomobject]@{ DocumentGUID = $staleGuid; Name = $DocumentName })
    }
    function Get-PWDocumentsByGUIDs {
        param([string[]]$DocumentGUIDs)
        $g = [string]$DocumentGUIDs[0]
        if ($g -eq $liveGuid) {
            return @([pscustomobject]@{ Name = $qcName; DocumentGUID = $liveGuid })
        }
        if ($g -eq $chkGuid) {
            return @([pscustomobject]@{ Name = $chkName; DocumentGUID = $chkGuid })
        }
        return @()
    }

    function Invoke-QCDatabaseQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})

        if ($Sql -match 'sheet_package_qc_pdfs') {
            if ($Parameters -and $Parameters.qcPdfName -eq '080J082001ab001-prod.pdf') {
                $table = New-Object System.Data.DataTable
                [void]$table.Columns.Add('document_guid', [string])
                [void]$table.Rows.Add($liveGuid)
                return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
            }
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
        }
        if ($Sql -match 'FROM sheet_packages' -and $Sql -match 'qc_pdf_guid') {
            if ($Parameters -and $Parameters.qcPdfName -eq '080J082001ab001-prod.pdf') {
                $table = New-Object System.Data.DataTable
                [void]$table.Columns.Add('lane_guid', [string])
                [void]$table.Rows.Add($liveGuid)
                return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
            }
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
        }
        if ($Sql -match 'sheet_documents') {
            $table = New-Object System.Data.DataTable
            [void]$table.Columns.Add('document_guid', [string])
            [void]$table.Rows.Add($staleGuid)
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $table }
        }
        if ($Sql -match 'FROM sheet_index' -and $Sql -match 'document_guid') {
            return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
        }
        return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
    }

    $resolved = _QCN-ResolveLiveQcPdfDocumentGuidResult -Config $config -FolderPath $folder `
        -QcPdfName '080J082001ab001-prod.pdf' -SheetStem '080J082001ab001' -HintGuid $staleGuid `
        -Event @{ qcProcessType = 'production' }
    Assert-Eq $resolved.documentGuid $liveGuid 'lane QC PDF GUID should win over stale sheet_documents, PW search, hint, and sheet_index'
    Assert-Eq $resolved.resolutionSource 'sheet_package_qc_pdfs' 'resolution source should identify lane table'

    $script:pwSearchFolders = @()
    $chkResolved = _QCN-ResolveLiveQcPdfDocumentGuidResult -Config $config -FolderPath $folder `
        -QcPdfName $chkName -SheetStem '0818000063ea501' -HintGuid $staleGuid `
        -Event @{ qcProcessType = 'check' }
    Assert-Eq $chkResolved.documentGuid $chkGuid 'check lane PDF should resolve via PW search when DB index is empty'
    Assert-Eq $chkResolved.resolutionSource 'pw_search' 'check lane should use pw_search before stale index'
    Assert-True ($script:pwSearchFolders -contains $apiFolder) 'PW search should use cmdlet folder path without Documents prefix'

    function Get-PWDocumentsInFolder {
        param([string]$FolderPath)
        if ($FolderPath -eq $apiFolder) {
            return @([pscustomobject]@{ DocumentGUID = $mixedChkGuid; Name = '080J082001ab001-chk.pdf' })
        }
        return @()
    }
    $caseResolved = _QCN-TryResolveQcPdfGuidFromPwSearch -Config $config -FolderPath $folder -QcPdfName $lowerChkName
    Assert-Eq $caseResolved $mixedChkGuid 'PW folder scan should resolve lane PDF with case-insensitive name match'

    $trustedUrl = _QCN-BuildPwDocumentLinkUrl -DocumentGuid $chkGuid -Settings @{
        email = @{ pwLinkBaseUrl = 'https://example.test/pwlink/'; pwLinkApp = 'pwe' }
    } -Config @{ projectWise = @{ datasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03' } }
    Assert-True ($trustedUrl -match 'objectId=' + [regex]::Escape($chkGuid)) 'trusted lane GUID should build pwlink URL'

    $laneTarget = _QCN-ResolveQcPdfNotificationTarget -Document $null -Config $config -Job $null `
        -DocumentName $chkName -DocumentGuid $staleGuid `
        -DocumentPath ($folder + '\' + $chkName)
    Assert-Eq $laneTarget.documentGuid $chkGuid 'lane PDF notification target should replace stem GUID via PW search'
    Assert-Eq $laneTarget.resolutionSource 'lane_pdf_pw_search' 'lane PDF target should identify pw_search resolution'

    $url = 'https://example.test/pwlink?objectId=0f9c6ba8-a5e1-40ed-b2a3-2b906cb4f38b&objectType=doc&datasource=x&app=pwe'
    Assert-Eq (_QCN-ExtractPwLinkDocumentGuid -Url $url) $liveGuid 'pwlink objectId should be extracted from URL'

    $legacyDoc = @{ QC_Review_Type = '0' }
    $prodJob = @{ metadata = @{ qcProcessType = 'production' } }
    $prodReview = _QCN-ResolveNotificationReviewType -Document $legacyDoc -Settings @{} -Config @{} -Job $prodJob `
        -DocumentName '080J082001ab001-prod.pdf'
    Assert-Eq $prodReview 'Production' 'review type should use qcProcessType instead of deprecated QC_Review_Type 0'

    $laneOnlyReview = _QCN-ResolveNotificationReviewType -Document $legacyDoc -Settings @{} -Config @{} -Job $null `
        -DocumentName '080J082001ab001-prod.pdf'
    Assert-Eq $laneOnlyReview 'Production' 'lane PDF suffix should infer Production review type label'
}

Write-Host 'OK: QC notification GUID resolution tests passed.' -ForegroundColor Green
