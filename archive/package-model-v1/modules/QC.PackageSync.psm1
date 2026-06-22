# Archived v1: resolve production modules from repo modules/
$script:_QCPkgV1RepoModules = Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))) 'modules'
# QC.PackageSync.psm1
# Responsibility: Apply package-level attribute synchronization without blindly copying metadata.

Import-Module (Join-Path $script:_QCPkgV1RepoModules 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.PackageResolver.psm1') -Force -Global
Import-Module (Join-Path $PSScriptRoot 'QC.AttributePolicy.psm1') -Force -Global

function _QCPS-Get([object]$Object,[string[]]$Names){ foreach($n in @($Names)){ try{ if($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n){ return $Object.$n }}catch{}; if($Object -is [hashtable] -and $Object.ContainsKey($n)){ return $Object[$n] } }; return $null }
function _QCPS-Attrs([object]$Document){ $a=_QCPS-Get $Document @('Attributes','attributes','EnvironmentAttributes'); if($a -is [hashtable]){return $a}; $h=@{}; if($a -and $a.PSObject.Properties){foreach($p in $a.PSObject.Properties){$h[$p.Name]=$p.Value}}; return $h }
function _QCPS-Log([string]$Event,[string]$Level,[string]$Message,[hashtable]$Data){ try { if(Get-Command Write-QCJsonLog -ErrorAction SilentlyContinue){Write-QCJsonLog -Level $Level -Code $Event -Message $Message -Data $Data|Out-Null} elseif(Get-Command Write-QCLog -ErrorAction SilentlyContinue){$Data.event=$Event; Write-QCLog -Level $Level -Message $Message -Data $Data|Out-Null} } catch {} }

function Sync-QCPackageAttributes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Package,
        [hashtable]$Config=@{},
        [hashtable]$AutomationValues=@{},
        [switch]$DryRun
    )
    $policy=(Get-QCPackageAttributePolicy -Config $Config).Data
    $canon=Get-QCPackageCanonicalDocument -Package $Package -Config $Config
    if(-not $canon.IsSuccess){ return $canon }
    $canonicalDoc=$canon.Data.document; $canonicalRole=[string]$canon.Data.role
    $canonicalAttrs=_QCPS-Attrs $canonicalDoc
    $roleToDoc=@{Dgn=$Package.DgnDocument;ProductionPdf=$Package.PdfDocument;QcPdf=$Package.QcPdfDocument}
    $writes=[System.Collections.Generic.List[object]]::new(); $conflicts=[System.Collections.Generic.List[object]]::new()

    if([bool]$policy.syncUserAttributesFromNonCanonical){
        foreach($role in @('Dgn','QcPdf')){
            if($role -eq $canonicalRole){ continue }
            $doc=$roleToDoc[$role]; if(-not $doc){ continue }
            $attrs=_QCPS-Attrs $doc; $toWrite=@{}
            foreach($name in @($policy.userOwnedAttributes)){
                if($attrs.ContainsKey($name) -and $attrs[$name] -and ((-not $canonicalAttrs.ContainsKey($name)) -or -not $canonicalAttrs[$name])){ $toWrite[$name]=$attrs[$name] }
                elseif($attrs.ContainsKey($name) -and $canonicalAttrs.ContainsKey($name) -and $attrs[$name] -and $canonicalAttrs[$name] -and ([string]$attrs[$name] -ne [string]$canonicalAttrs[$name])){ $conflicts.Add([pscustomobject]@{ Attribute=$name; CanonicalValue=$canonicalAttrs[$name]; NonCanonicalRole=$role; NonCanonicalValue=$attrs[$name] }) | Out-Null }
            }
            if($toWrite.Count -gt 0){
                $writes.Add([pscustomobject]@{ Role=$canonicalRole; DocumentGuid=(_QCPS-Get $canonicalDoc @('Guid','DocumentGuid','ObjectGuid','guid')); Ownership='UserOwned'; Attributes=$toWrite; SourceRole=$role; DryRun=[bool]$DryRun }) | Out-Null
                if(-not $DryRun){ Update-PWDocumentAttributes -InputDocument $canonicalDoc -Attributes $toWrite | Out-Null }
            }
        }
    }

    foreach($role in @($policy.writeAutomationStatusTo)){
        $doc=$roleToDoc[[string]$role]; if(-not $doc -or $AutomationValues.Count -eq 0){ continue }
        $toWrite=@{}
        foreach($name in @($policy.automationOwnedAttributes)){ if($AutomationValues.ContainsKey($name)){ $toWrite[$name]=$AutomationValues[$name] } }
        if($toWrite.Count -gt 0){
            $writes.Add([pscustomobject]@{ Role=$role; DocumentGuid=(_QCPS-Get $doc @('Guid','DocumentGuid','ObjectGuid','guid')); Ownership='AutomationOwned'; Attributes=$toWrite; DryRun=[bool]$DryRun }) | Out-Null
            if(-not $DryRun){ Update-PWDocumentAttributes -InputDocument $doc -Attributes $toWrite | Out-Null }
        }
    }
    if($conflicts.Count -gt 0){ _QCPS-Log -Event 'PACKAGE_ATTRIBUTE_CONFLICT' -Level 'Warning' -Message 'QC package attribute conflicts detected.' -Data @{ packageId=$Package.PackageId; conflicts=@($conflicts) } }
    _QCPS-Log -Event 'PACKAGE_ATTRIBUTES_SYNCED' -Level 'Information' -Message 'QC package attribute sync planned or completed.' -Data @{ packageId=$Package.PackageId; dryRun=[bool]$DryRun; writes=@($writes); conflicts=@($conflicts) }
    return New-QCSuccessResult -Code 'PACKAGE_ATTRIBUTES_SYNCED' -Message 'QC package attribute sync planned or completed.' -Data @{ canonicalRole=$canonicalRole; writes=@($writes); conflicts=@($conflicts); dryRun=[bool]$DryRun }
}
Export-ModuleMember -Function Sync-QCPackageAttributes
