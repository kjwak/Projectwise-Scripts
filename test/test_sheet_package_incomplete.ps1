# Incomplete packages (PDF only, no DGN/QC PDF) still create a sheet_packages row.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$script:lastMergeRole = $null

InModuleScope -ModuleName Core.Database {
    function _QDB-IsEnabled { param([hashtable]$Config) return $true }
    function _QDB-NormalizeTelemetryPath { param([string]$Path) return ([string]$Path).Trim().ToLowerInvariant() }

    function Invoke-QCDatabaseNonQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ rowsAffected = 1 }
    }

    function Invoke-QCDatabaseScalar {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        if ($Parameters.ContainsKey('documentRole')) { $script:lastMergeRole = $Parameters.documentRole }
        return New-QCSuccessResult -Code 'DB_SCALAR_OK' -Message 'ok' -Data @{ value = [guid]::NewGuid() }
    }

    $config = @{ database = @{ enabled = $true; connectionString = 'x' } }
    $resolved = Resolve-SheetPackageFromDocument -DocumentGuid 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' `
        -DocumentName '080J082001ab002.pdf' -FolderPath 'documents\proj\cadd\sheets'
    Assert-True $resolved.isSheetPackageMember 'orphan sheet PDF is still a package member'

    $pkg = Ensure-SheetPackage -Config $config -FolderPath $resolved.folderPath -SheetStem $resolved.sheetStem `
        -DocumentRole $resolved.documentRole -DocumentGuid $resolved.documentGuid -DocumentName $resolved.documentName `
        -PwStateName 'QC Received'
    Assert-True $pkg.IsSuccess 'Ensure-SheetPackage should succeed for PDF-only package'
    Assert-Eq $script:lastMergeRole 'sheet_pdf' 'MERGE should use sheet_pdf role'
    Assert-True ($null -ne $pkg.Data.sheetPackageId) 'sheet_package_id should be assigned'
}

Write-Host 'OK: incomplete sheet package tests passed.' -ForegroundColor Green
