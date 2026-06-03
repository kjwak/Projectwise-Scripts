# QC.StatePolicy.psm1
# Responsibility: Package-level workflow state precedence, conflict detection, and idempotent state planning.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.PackageResolver.psm1') -Force

function _QCSP-Get([object]$Object,[string[]]$Names){ foreach($n in @($Names)){ try{ if($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n){ return $Object.$n }}catch{}; if($Object -is [hashtable] -and $Object.ContainsKey($n)){ return $Object[$n] } }; return $null }
function _QCSP-State([object]$Doc){ return [string](_QCSP-Get $Doc @('WorkflowState','StateName','CurrentState','DocumentState','state')) }
function _QCSP-Log([string]$Event,[string]$Level,[string]$Message,[hashtable]$Data){ try { if(Get-Command Write-QCJsonLog -ErrorAction SilentlyContinue){Write-QCJsonLog -Level $Level -Code $Event -Message $Message -Data $Data|Out-Null} elseif(Get-Command Write-QCLog -ErrorAction SilentlyContinue){$Data.event=$Event; Write-QCLog -Level $Level -Message $Message -Data $Data|Out-Null} } catch {} }

function Resolve-QCPackageState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Package,[hashtable]$Config=@{})
    $s = Get-QCPackageSettings -Config $Config
    $states=@(); foreach($pair in @(@('Dgn',$Package.DgnDocument),@('ProductionPdf',$Package.PdfDocument),@('QcPdf',$Package.QcPdfDocument))){ $st=_QCSP-State $pair[1]; if($st){ $states += [pscustomobject]@{ Role=$pair[0]; State=$st } } }
    if($states.Count -eq 0){ return New-QCSuccessResult -Code 'PACKAGE_STATE_EMPTY' -Message 'No package document states were present.' -Data @{ state=$null; states=@(); conflict=$false } }
    $selected=$states[0].State; $rank=999999
    foreach($item in $states){ $i=[array]::IndexOf(@($s.statePrecedence), $item.State); if($i -lt 0){ $i=9999 }; if($i -lt $rank){ $rank=$i; $selected=$item.State } }
    $distinct=@($states.State | Select-Object -Unique); $conflict=($distinct.Count -gt 1)
    if($conflict){ _QCSP-Log -Event 'PACKAGE_STATE_CONFLICT' -Level 'Warning' -Message 'QC package documents have conflicting workflow states.' -Data @{ packageId=$Package.PackageId; selectedState=$selected; states=@($states) } }
    return New-QCSuccessResult -Code 'PACKAGE_STATE_RESOLVED' -Message 'QC package state resolved by configured precedence.' -Data @{ state=$selected; states=@($states); conflict=$conflict; precedence=@($s.statePrecedence) }
}
function Test-QCPackageStateConflict { [CmdletBinding()] param([Parameter(Mandatory)][hashtable]$Package,[hashtable]$Config=@{}) $r=Resolve-QCPackageState -Package $Package -Config $Config; return New-QCSuccessResult -Code 'PACKAGE_STATE_CONFLICT_TESTED' -Message 'QC package state conflict check completed.' -Data @{ hasConflict=[bool]$r.Data.conflict; resolvedState=$r.Data.state; states=$r.Data.states } }
function Set-QCPackageState {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][hashtable]$Package,[Parameter(Mandatory)][string]$StateName,[hashtable]$Config=@{},[string[]]$Roles=@('ProductionPdf','QcPdf'),[switch]$DryRun)
    if ([string]::IsNullOrWhiteSpace($StateName)) {
        return New-QCFailureResult -Code 'PACKAGE_STATE_EMPTY_TARGET' -Message 'Target workflow state is empty; package state sync was not performed.' -Data @{ packageId = $Package.PackageId }
    }
    $roleToDoc=@{Dgn=$Package.DgnDocument;ProductionPdf=$Package.PdfDocument;QcPdf=$Package.QcPdfDocument}; $actions=[System.Collections.Generic.List[object]]::new()
    foreach($role in @($Roles)){ $doc=$roleToDoc[$role]; if(-not $doc){ continue }; $cur=_QCSP-State $doc; $guid=[string](_QCSP-Get $doc @('Guid','DocumentGuid','ObjectGuid','guid'))
        $planned = -not ($cur -ieq $StateName); $changed=$false
        if($planned -and -not $DryRun){ if(Get-Command Set-PWDocumentState -ErrorAction SilentlyContinue){ Set-PWDocumentState -InputDocument $doc -StateName $StateName | Out-Null; $changed=$true } else { return New-QCFailureResult -Code 'PACKAGE_STATE_CMDLET_MISSING' -Message 'Set-PWDocumentState is not available.' -Data @{ role=$role; documentGuid=$guid } } }
        $actions.Add([pscustomobject]@{ Role=$role; DocumentGuid=$guid; CurrentState=$cur; TargetState=$StateName; Planned=$planned; Changed=$changed; DryRun=[bool]$DryRun }) | Out-Null
    }
    _QCSP-Log -Event 'PACKAGE_STATE_SYNCED' -Level 'Information' -Message 'QC package state sync planned or completed.' -Data @{ packageId=$Package.PackageId; state=$StateName; dryRun=[bool]$DryRun; actions=@($actions) }
    return New-QCSuccessResult -Code 'PACKAGE_STATE_SYNCED' -Message 'QC package state sync planned or completed.' -Data @{ actions=@($actions); dryRun=[bool]$DryRun }
}
Export-ModuleMember -Function Resolve-QCPackageState,Test-QCPackageStateConflict,Set-QCPackageState
