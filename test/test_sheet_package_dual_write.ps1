# Dual-write: Write-QCSheetIndex triggers sheet package upsert (mocked SQL).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

InModuleScope -ModuleName Core.Database {
    function Test-QCSheetIndexFolderPath { param([string]$FolderPath) return $true }
    function _QDB-IsEnabled { param([hashtable]$Config) return $true }
    function _QDB-NormalizeTelemetryPath { param([string]$Path) return ([string]$Path).Trim() }
    function Write-QCJsonLog { param([hashtable]$Data) }

    function Invoke-QCDatabaseNonQuery {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        return New-QCSuccessResult -Code 'DB_OK' -Message 'ok' -Data @{ rowsAffected = 1 }
    }

    function Invoke-QCDatabaseScalar {
        param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters = @{})
        return New-QCSuccessResult -Code 'DB_SCALAR_OK' -Message 'ok' -Data @{ value = [guid]::NewGuid() }
    }

    function Invoke-QCDatabaseQuery {
        return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message 'ok' -Data @{ table = $null }
    }

    $config = @{ database = @{ enabled = $true; connectionString = 'x' } }
    $res = Write-QCSheetIndex -Config $config -DocumentGuid 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' `
        -DocumentName '080J082001ab001.dgn' -FolderPath 'Documents\Proj\CADD\Sheets' `
        -PwStateName 'QC Initiated' -DesignerEmail 'designer@example.com'

    Assert-True $res.IsSuccess "Write-QCSheetIndex should succeed ($($res.Code): $($res.Message))"
    Assert-True ($null -ne $res.Data.sheetPackageId) 'sheetPackageId should be returned after dual-write'
}

Write-Host 'OK: sheet package dual-write tests passed.' -ForegroundColor Green
