# Archived v1: resolve production modules from repo modules/
$script:_QCPkgV1RepoModules = Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))) 'modules'
# QC.Package.Database.psm1
# Responsibility: Additive SQL/cache extension points for QC package relationships.

Import-Module (Join-Path $script:_QCPkgV1RepoModules 'Core.Results.psm1') -Force
Import-Module (Join-Path $script:_QCPkgV1RepoModules 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue

function _QCPD-Get([object]$Object,[string[]]$Names){ foreach($n in @($Names)){ try{ if($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n){ return $Object.$n }}catch{}; if($Object -is [hashtable] -and $Object.ContainsKey($n)){ return $Object[$n] } }; return $null }
function _QCPD-Guid([object]$Doc){ return [string](_QCPD-Get $Doc @('Guid','DocumentGuid','ObjectGuid','guid')) }
function _QCPD-Log([string]$Event,[string]$Level,[string]$Message,[hashtable]$Data){ try { if(Get-Command Write-QCJsonLog -ErrorAction SilentlyContinue){Write-QCJsonLog -Level $Level -Code $Event -Message $Message -Data $Data|Out-Null} elseif(Get-Command Write-QCLog -ErrorAction SilentlyContinue){$Data.event=$Event; Write-QCLog -Level $Level -Message $Message -Data $Data|Out-Null} } catch {} }

function Write-QCPackageCache {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Package,[hashtable]$Config=@{},[switch]$DryRun)
    $row=@{
        PackageId=$Package.PackageId; DgnGuid=(_QCPD-Guid $Package.DgnDocument); PdfGuid=(_QCPD-Guid $Package.PdfDocument); QcPdfGuid=(_QCPD-Guid $Package.QcPdfDocument)
        PackageKey=$Package.PackageKey; LastResolvedUtc=$(if(Get-Command Get-QCTimestamp -ErrorAction SilentlyContinue){Get-QCTimestamp}else{[DateTime]::UtcNow.ToString('o')})
        LastCanonicalGuid=$null; LastKnownState=$Package.WorkflowState; HasConflict=(@($Package.Conflicts).Count -gt 0)
    }
    if($DryRun){ return New-QCSuccessResult -Code 'PACKAGE_CACHE_DRY_RUN' -Message 'Dry-run: package cache row was not written.' -Data @{ row=$row; dryRun=$true } }
    if(Get-Command Invoke-QCSqlNonQuery -ErrorAction SilentlyContinue){
        Invoke-QCSqlNonQuery -Config $Config -Query 'MERGE dbo.QCPackageCache WITH additive schema contract' -Parameters $row | Out-Null
        return New-QCSuccessResult -Code 'PACKAGE_CACHE_WRITTEN' -Message 'Package cache row written.' -Data @{ row=$row; dryRun=$false }
    }
    return New-QCFailureResult -Code 'PACKAGE_CACHE_WRITER_MISSING' -Message 'No SQL non-query helper is available; package cache was not written.' -Data @{ row=$row }
}

function Get-QCPackageByDocumentGuid {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DocumentGuid,[hashtable]$Config=@{})
    if(Get-Command Invoke-QCSqlQuery -ErrorAction SilentlyContinue){
        $rows=Invoke-QCSqlQuery -Config $Config -Query 'SELECT * FROM dbo.QCPackageCache WHERE @DocumentGuid IN (DgnGuid,PdfGuid,QcPdfGuid)' -Parameters @{ DocumentGuid=$DocumentGuid }
        return New-QCSuccessResult -Code 'PACKAGE_CACHE_FOUND' -Message 'Package cache lookup completed.' -Data @{ rows=@($rows); documentGuid=$DocumentGuid }
    }
    return New-QCSuccessResult -Code 'PACKAGE_CACHE_LOOKUP_UNAVAILABLE' -Message 'No SQL query helper is available; package cache lookup skipped.' -Data @{ rows=@(); documentGuid=$DocumentGuid }
}
Export-ModuleMember -Function Write-QCPackageCache,Get-QCPackageByDocumentGuid
