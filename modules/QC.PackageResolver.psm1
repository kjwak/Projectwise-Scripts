# QC.PackageResolver.psm1
# Responsibility: Resolve related ProjectWise documents into a package-level QC model.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'Core.Logging.psm1') -Force -ErrorAction SilentlyContinue

function _QCPR-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value.PSObject -and $Value.PSObject.Properties) { $h = @{}; foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }; return $h }
    return $null
}
function _QCPR-Get([object]$Object, [string[]]$Names) { foreach ($n in @($Names)) { try { if ($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n) { return $Object.$n } } catch { }; if ($Object -is [hashtable] -and $Object.ContainsKey($n)) { return $Object[$n] } }; return $null }
function _QCPR-IsBlank([object]$Value) { return ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) }
function _QCPR-Log([string]$Event, [string]$Level, [string]$Message, [hashtable]$Data) { try { if (Get-Command Write-QCJsonLog -ErrorAction SilentlyContinue) { Write-QCJsonLog -Level $Level -Code $Event -Message $Message -Data $Data | Out-Null } elseif (Get-Command Write-QCLog -ErrorAction SilentlyContinue) { $d=@{}; if($Data){foreach($k in $Data.Keys){$d[$k]=$Data[$k]}}; $d.event=$Event; Write-QCLog -Level $Level -Message $Message -Data $d | Out-Null } } catch { } }

function Get-QCPackageSettings {
    [CmdletBinding()]
    param([hashtable]$Config)
    $defaults = @{
        enabled = $true
        canonicalDocumentRole = 'ProductionPdf'
        fallbackOrder = @('ProductionPdf','QcPdf','Dgn')
        namingRules = @{ dgnExtensions=@('.dgn'); productionPdfExtensions=@('.pdf'); qcPdfSuffixes=@('_QC','-QC','_QC_HISTORY') }
        statePrecedence = @('Error Needs Attention','QC Initiated','QC Finalizing','Ready for QC','Corrections Received','Redlines Received','In Production')
        userOwnedAttributes = @('QC_Review_Type','QC_Designer_Email','QC_Reviewer_Email','QC_Checker_Email','QC_Originator','QC_Due_Date')
        automationOwnedAttributes = @('QC_Prepend_Status','QC_Last_Automation_Run','QC_Last_Email_Sent','QC_Automation_JobId','QC_Error_Message','QC_PackageId','QC_Source_Dgn_Guid','QC_Production_Pdf_Guid','QC_QcPdf_Guid')
        syncUserAttributesFromNonCanonical = $true
        nonCanonicalUserEditBehavior = 'CopyOnceThenNormalize'
        writeAutomationStatusTo = @('ProductionPdf','QcPdf')
    }
    $raw = $null
    if ($Config -and $Config.ContainsKey('qcPackage')) { $raw = _QCPR-ToHashtable $Config.qcPackage }
    if (-not $raw) { return $defaults }
    $settings = @{}
    foreach ($k in $defaults.Keys) { $settings[$k] = $defaults[$k] }
    foreach ($k in $raw.Keys) { if ($null -ne $raw[$k]) { $settings[$k] = $raw[$k] } }
    return $settings
}

function Get-QCPackageDocumentRole {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Document,[hashtable]$Config)
    $settings = Get-QCPackageSettings -Config $Config
    $name = [string](_QCPR-Get $Document @('Name','DocumentName','FileName','name','fileName'))
    $ext = ([System.IO.Path]::GetExtension($name)).ToLowerInvariant()
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
    if ($ext -in @($settings.namingRules.dgnExtensions | ForEach-Object { ([string]$_).ToLowerInvariant() })) { return 'Dgn' }
    if ($ext -in @($settings.namingRules.productionPdfExtensions | ForEach-Object { ([string]$_).ToLowerInvariant() })) {
        foreach ($s in @($settings.namingRules.qcPdfSuffixes)) { if ($stem.EndsWith([string]$s, [System.StringComparison]::OrdinalIgnoreCase)) { return 'QcPdf' } }
        return 'ProductionPdf'
    }
    return 'Unknown'
}

function Get-QCPackageDocumentKey {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Document,[hashtable]$Config)
    $settings = Get-QCPackageSettings -Config $Config
    $name = [string](_QCPR-Get $Document @('Name','DocumentName','FileName','name','fileName'))
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
    foreach ($s in @($settings.namingRules.qcPdfSuffixes)) { if ($stem.EndsWith([string]$s, [System.StringComparison]::OrdinalIgnoreCase)) { $stem = $stem.Substring(0, $stem.Length - ([string]$s).Length); break } }
    return $stem
}

function _QCPR-NewPackageId([string]$PackageKey, [string]$Folder) {
    $text = (($Folder -as [string]) + '|' + ($PackageKey -as [string])).ToLowerInvariant()
    $sha = [System.Security.Cryptography.SHA256]::Create(); try { $hex = -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)) | ForEach-Object { $_.ToString('x2') }) } finally { $sha.Dispose() }
    return 'qcpkg_' + $hex.Substring(0,24)
}

function Resolve-QCPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Document,
        [hashtable]$Config = @{},
        [object[]]$CandidateDocuments = @(),
        [hashtable]$CachedPackage,
        [switch]$DryRun
    )
    $settings = Get-QCPackageSettings -Config $Config
    if (-not [bool]$settings.enabled) { return New-QCFailureResult -Code 'PACKAGE_DISABLED' -Message 'QC Package resolution is disabled.' -Data @{ dryRun=[bool]$DryRun } }
    $docGuid = [string](_QCPR-Get $Document @('Guid','DocumentGuid','ObjectGuid','guid','documentGuid'))
    $folder = [string](_QCPR-Get $Document @('FolderPath','ProjectWiseFolder','Path','folderPath','sourceFolder'))
    $role = Get-QCPackageDocumentRole -Document $Document -Config $Config
    $key = Get-QCPackageDocumentKey -Document $Document -Config $Config
    $docs = @($CandidateDocuments) + @($Document)
    $siblings = @{ Dgn=$null; ProductionPdf=$null; QcPdf=$null }
    foreach ($d in $docs) {
        $r = Get-QCPackageDocumentRole -Document $d -Config $Config
        if (-not $siblings.ContainsKey($r)) { continue }
        $dk = Get-QCPackageDocumentKey -Document $d -Config $Config
        $df = [string](_QCPR-Get $d @('FolderPath','ProjectWiseFolder','Path','folderPath','sourceFolder'))
        if ($dk -ieq $key -and ((_QCPR-IsBlank $folder) -or (_QCPR-IsBlank $df) -or $df -ieq $folder)) { $siblings[$r] = $d }
    }
    if ($CachedPackage) {
        foreach ($r in @('Dgn','ProductionPdf','QcPdf')) { if (-not $siblings[$r] -and $CachedPackage.ContainsKey(($r + 'Document'))) { $siblings[$r] = $CachedPackage[($r + 'Document')] } }
    }
    $pkgId = [string](_QCPR-Get $Document @('QC_PackageId','PackageId','packageId'))
    if (_QCPR-IsBlank $pkgId -and $CachedPackage -and $CachedPackage.ContainsKey('PackageId')) { $pkgId = [string]$CachedPackage.PackageId }
    if (_QCPR-IsBlank $pkgId) { $pkgId = _QCPR-NewPackageId -PackageKey $key -Folder $folder }
    $package = @{
        PackageId = $pkgId; PackageKey = $key; DocumentKey = $key; BaseName = $key; SheetId = $key; PackageRootFolder = $folder
        TriggerDocumentGuid = $docGuid; TriggerDocumentRole = $role
        DgnDocument = $siblings.Dgn; PdfDocument = $siblings.ProductionPdf; QcPdfDocument = $siblings.QcPdf
        WorkflowState = $null; ReviewType = $null; Conflicts = @(); Warnings = @(); DryRun = [bool]$DryRun
        ResolutionStrategies = @('ExplicitPackageAttribute','GuidCache','SqlPackageCache','FolderProximity','NamingRules')
    }
    _QCPR-Log -Event 'PACKAGE_RESOLVED' -Level 'Information' -Message 'QC package resolved.' -Data @{ packageId=$pkgId; packageKey=$key; triggerDocumentGuid=$docGuid; triggerDocumentRole=$role; dryRun=[bool]$DryRun }
    return New-QCSuccessResult -Code 'PACKAGE_RESOLVED' -Message 'QC package resolved.' -Data @{ package=$package; settings=$settings }
}

function Get-QCPackageCanonicalDocument {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Package,[hashtable]$Config = @{})
    $settings = Get-QCPackageSettings -Config $Config
    $roleToProp = @{ ProductionPdf='PdfDocument'; QcPdf='QcPdfDocument'; Dgn='DgnDocument' }
    $warnings = [System.Collections.Generic.List[string]]::new()
    foreach ($role in @($settings.fallbackOrder)) {
        $prop = $roleToProp[[string]$role]
        if ($prop -and $Package.ContainsKey($prop) -and $Package[$prop]) {
            if ($role -ne $settings.canonicalDocumentRole) { $warnings.Add("Preferred canonical document '$($settings.canonicalDocumentRole)' was unavailable; fell back to '$role'.") | Out-Null }
            _QCPR-Log -Event 'PACKAGE_CANONICAL_SELECTED' -Level 'Information' -Message 'QC package canonical document selected.' -Data @{ packageId=$Package.PackageId; role=$role; warnings=@($warnings) }
            return New-QCSuccessResult -Code 'PACKAGE_CANONICAL_SELECTED' -Message 'QC package canonical document selected.' -Data @{ document=$Package[$prop]; role=$role; warnings=@($warnings) }
        }
    }
    _QCPR-Log -Event 'PACKAGE_RESOLVE_FAILED' -Level 'Warning' -Message 'No canonical document could be selected for QC package.' -Data @{ packageId=$Package.PackageId }
    return New-QCFailureResult -Code 'PACKAGE_CANONICAL_MISSING' -Message 'No configured canonical/fallback document exists in the package.' -Data @{ package=$Package }
}

Export-ModuleMember -Function Get-QCPackageSettings,Get-QCPackageDocumentRole,Get-QCPackageDocumentKey,Resolve-QCPackage,Get-QCPackageCanonicalDocument
