# QC process type attribute sync: legacy watcher path disabled by default; prepend-only stem reset.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

Assert-False (Test-QCLegacyReviewTypeAttributeSyncEnabled -Config @{}) 'legacy review/process type sync default off'

$legacyCfg = @{ QCProcess = @{ EnableLegacyReviewTypeAttributeSync = $true } }
Assert-True (Test-QCLegacyReviewTypeAttributeSyncEnabled -Config $legacyCfg) 'legacy flag enables attribute sync'

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) $script:lastLog = @{ Code = $Code; Data = $Data } }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Get-PWQcReviewTypeAttributeName { param($Config) return 'QC_Review_Type' }
    function Test-PWQcReviewTypeAttributesEnabled { param($Config, $FolderPath) return $true }
    function _PWD-GetSheetIndexQcReviewType { param($Config, $DocumentGuid) return 'Check' }
    function _PWD-NormalizeSheetIndexValue { param($Value) return ([string]$Value).Trim().ToLowerInvariant() }
    function _PWD-GetPwAttributeValue { param($PwAttributes, $ColumnName) return 'Check' }
    function _PWD-ResolvePwDocumentInFolder { param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid) return [pscustomobject]@{ Name = $DocumentName } }
    function _PWD-InvokeUpdatePWDocumentAttributes { param($Document, $Attributes, $Config) $script:attrWrite = $Attributes; return $true }
    function Get-PWAssociatedSheetSyncMembers {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TriggerSource)
        return @(@{ documentGuid = '1'; documentName = '0818000063ea500.pdf' })
    }
    function Get-PWDocumentAttributesByColumns { param($FolderPath, $DocumentName, $ColumnsToReturn) return @{ found = $true; attributes = @{} } }
    function Get-PWQcPdfLaneFromDocumentName { param($DocumentName) return $null }
    function Invoke-QCAuditWorkflowAttributeChangeTriggers { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $FieldChanges) }
    function Write-QCSheetIndex { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $WatchRoot, $Extension, $SourceType, $QcReviewType, $LastAuditEventAt, $SetOwnershipFromProjectWise) }

    $script:attrWrite = $null
    $script:lastLog = $null
    Sync-PWAssociatedSheetReviewTypeAttributes -Config @{} -DocumentGuid '1' -DocumentName '0818000063ea500.dgn' `
        -FolderPath 'Drawings\X' -CanonicalReviewType 'Production'
    Assert-True ($null -eq $script:attrWrite) 'default config must not write PW attributes from watcher sync'
    Assert-Eq $script:lastLog.Code 'QC_PROCESS_TYPE_SYNC_SKIPPED' 'logs skip when legacy sync disabled'
    Assert-Eq $script:lastLog.Data.reason 'legacy_sync_disabled' 'skip reason legacy_sync_disabled'

    $script:attrWrite = $null
    Sync-PWAssociatedSheetReviewTypeAttributes -Config @{} -DocumentGuid '1' -DocumentName '0818000063ea500.pdf' `
        -FolderPath 'Drawings\X' -CanonicalReviewType ''
    Assert-True ($null -eq $script:attrWrite) 'empty canonical must not write'
    Assert-Eq $script:lastLog.Code 'QC_PROCESS_TYPE_SYNC_SKIPPED' 'null canonical logs skip'

    $script:attrWrite = $null
    Sync-PWAssociatedSheetReviewTypeAttributes -Config $legacyCfg -DocumentGuid '1' -DocumentName '0818000063ea500.pdf' `
        -FolderPath 'Drawings\X' -CanonicalReviewType 'not-a-real-type'
    Assert-True ($null -eq $script:attrWrite) 'unknown canonical must not default to Production'
}

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-PWSheetStemFromDocumentName { param($DocumentName) return '0818000063ea500' }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Test-PWQcReviewTypeAttributesEnabled { param($Config, $FolderPath) return $true }
    function _PWD-NormalizeSheetIndexValue { param($Value) return ([string]$Value).Trim().ToLowerInvariant() }
    function _PWD-GetPwAttributeValue { param($PwAttributes, $ColumnName) return 'Check' }
    function _PWD-ResolvePwDocumentInFolder { param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid) return [pscustomobject]@{ DocumentGUID = 'stem'; Name = $DocumentName } }
    function Get-PWDocumentAttributesByColumns { param($FolderPath, $DocumentName, $ColumnsToReturn) return @{ found = $true; attributes = @{} } }
    function _PWD-InvokeUpdatePWDocumentAttributes { param($Document, $Attributes, $Config) $script:stemResetAttrs = $Attributes; return $true }
    function Write-QCSheetIndex { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $WatchRoot, $Extension, $SourceType, $QcReviewType, $LastAuditEventAt, $SetOwnershipFromProjectWise) }

    $script:stemResetAttrs = $null
    _PWD-SyncReferenceSheetProcessTypeAttributes -Config @{} -DocumentGuid 'stem' -DocumentName '0818000063ea500.pdf' `
        -FolderPath 'Drawings\X' -CanonicalProcessType 'production'
    Assert-True ($null -ne $script:stemResetAttrs) 'prepend stem reset writes QC_Process_Type'
    Assert-Eq $script:stemResetAttrs['QC_Process_Type'] 'Production' 'prepend reset formats QC_Process_Type display value'
    Assert-False ($script:stemResetAttrs.ContainsKey('QC_Review_Type')) 'prepend reset must not write QC_Review_Type'
}

Write-Host 'test_qc_process_type_attribute_sync: OK' -ForegroundColor Green
