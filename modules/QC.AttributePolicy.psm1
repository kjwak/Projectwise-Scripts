# QC.AttributePolicy.psm1
# Responsibility: Attribute ownership, validation, and allowlisted package metadata reads/writes.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.PackageResolver.psm1') -Force

function _QCAP-Get([object]$Object, [string[]]$Names) { foreach ($n in @($Names)) { try { if ($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n) { return $Object.$n } } catch { }; if ($Object -is [hashtable] -and $Object.ContainsKey($n)) { return $Object[$n] } }; return $null }
function _QCAP-Attrs([object]$Document) { $a = _QCAP-Get $Document @('Attributes','attributes','EnvironmentAttributes'); if ($a -is [hashtable]) { return $a }; $h=@{}; if($a -and $a.PSObject.Properties){foreach($p in $a.PSObject.Properties){$h[$p.Name]=$p.Value}}; return $h }
function _QCAP-IsBlank([object]$Value) { return ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) }
function _QCAP-IsEmailLike([string]$Value) { if (_QCAP-IsBlank $Value) { return $true }; return ($Value -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') }
function _QCAP-Log([string]$Event,[string]$Level,[string]$Message,[hashtable]$Data){ try { if(Get-Command Write-QCJsonLog -ErrorAction SilentlyContinue){Write-QCJsonLog -Level $Level -Code $Event -Message $Message -Data $Data|Out-Null} elseif(Get-Command Write-QCLog -ErrorAction SilentlyContinue){$Data.event=$Event; Write-QCLog -Level $Level -Message $Message -Data $Data|Out-Null} } catch {} }

function Get-QCPackageAttributePolicy {
    [CmdletBinding()]
    param([hashtable]$Config = @{})
    $s = Get-QCPackageSettings -Config $Config
    return New-QCSuccessResult -Code 'PACKAGE_ATTRIBUTE_POLICY' -Message 'QC package attribute policy resolved.' -Data @{
        userOwnedAttributes = @($s.userOwnedAttributes)
        automationOwnedAttributes = @($s.automationOwnedAttributes)
        writeAutomationStatusTo = @($s.writeAutomationStatusTo)
        syncUserAttributesFromNonCanonical = [bool]$s.syncUserAttributesFromNonCanonical
        nonCanonicalUserEditBehavior = [string]$s.nonCanonicalUserEditBehavior
    }
}
function Get-QCPackageUserAttributes { [CmdletBinding()] param([hashtable]$Config=@{}) $p=Get-QCPackageAttributePolicy -Config $Config; return $p.Data.userOwnedAttributes }
function Get-QCPackageAutomationAttributes { [CmdletBinding()] param([hashtable]$Config=@{}) $p=Get-QCPackageAttributePolicy -Config $Config; return $p.Data.automationOwnedAttributes }

function Get-QCPackageUserAttributeValues {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Document,[hashtable]$Config=@{})
    $attrs = _QCAP-Attrs $Document; $out=@{}
    foreach($name in @(Get-QCPackageUserAttributes -Config $Config)){ if($attrs.ContainsKey($name)){ $out[$name]=$attrs[$name] } }
    return $out
}

function Test-QCPackageAttributeValidity {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Document,[hashtable]$Config=@{},[string[]]$AttributeNames)
    if(-not $AttributeNames){ $AttributeNames = @(Get-QCPackageUserAttributes -Config $Config) + @(Get-QCPackageAutomationAttributes -Config $Config) }
    $warnings=[System.Collections.Generic.List[string]]::new(); $errors=[System.Collections.Generic.List[string]]::new()
    $envName = [string](_QCAP-Get $Document @('Environment','EnvironmentName','environment'))
    $available = $null
    if(Get-Command Get-PWEnvironmentColumns -ErrorAction SilentlyContinue){ try { $cols = Get-PWEnvironmentColumns -EnvironmentName $envName; $available = @($cols | ForEach-Object { [string](_QCAP-Get $_ @('Name','ColumnName','Column','name')) } | Where-Object { $_ }) } catch { $warnings.Add("Unable to discover ProjectWise environment columns for '$envName': $($_.Exception.Message)") | Out-Null } }
    elseif(-not (_QCAP-IsBlank $envName)) { $warnings.Add('Get-PWEnvironmentColumns is not available; attribute existence could not be verified against ProjectWise.') | Out-Null }
    if($available){ foreach($n in @($AttributeNames)){ if($available -notcontains $n){ $errors.Add("Missing environment attribute '$n' in '$envName'.") | Out-Null } } }
    $attrs = _QCAP-Attrs $Document
    foreach($n in @($AttributeNames | Where-Object { $_ -match '(?i)(email)$' })){ if($attrs.ContainsKey($n) -and -not (_QCAP-IsEmailLike ([string]$attrs[$n]))){ $errors.Add("Invalid email value for '$n'.") | Out-Null } }
    $ok = ($errors.Count -eq 0)
    if($ok){ return New-QCSuccessResult -Code 'PACKAGE_ATTRIBUTES_VALID' -Message 'QC package attributes are valid.' -Data @{ warnings=@($warnings); checked=@($AttributeNames); environment=$envName } }
    return New-QCFailureResult -Code 'PACKAGE_ATTRIBUTES_INVALID' -Message 'QC package attributes failed validation.' -Data @{ errors=@($errors); warnings=@($warnings); checked=@($AttributeNames); environment=$envName }
}

Export-ModuleMember -Function Get-QCPackageAttributePolicy,Get-QCPackageUserAttributes,Get-QCPackageAutomationAttributes,Get-QCPackageUserAttributeValues,Test-QCPackageAttributeValidity
