# Unit checks for audit folder GUID cache warm helpers and gate pass-through.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

# Discovery path extraction (hashtable bug fix)
$ht = @{ FolderPath = 'Documents\Proj\CADD\Sheets'; OneLevelDeep = $true }
_Assert ((Resolve-QCAuditSheetsDiscoveryFolderPath -Entry $ht) -eq 'Documents\Proj\CADD\Sheets') 'hashtable FolderPath'
_Assert ((Resolve-QCAuditSheetsDiscoveryFolderPath -Entry 'Documents\Proj\CADD\Sheets') -eq 'Documents\Proj\CADD\Sheets') 'string path'

# Exclusion filter
$seg = 'Documents\AZDOT 2024\Project\CADD\Sheets\Seg_1'
$refOrd = 'Documents\AZDOT 2024\Project\CADD\Ref-ORD'
$root = 'Documents\AZDOT 2024'
_Assert (-not (Test-QCAuditCacheWarmFolderPathExcluded -FolderPath $seg)) 'Seg_1 not excluded'
_Assert (-not (Test-QCAuditCacheWarmFolderPathExcluded -FolderPath $root)) 'watch root not excluded'
_Assert (Test-QCAuditCacheWarmFolderPathExcluded -FolderPath $refOrd) 'Ref-ORD excluded'
_Assert (Test-QCAuditCacheWarmFolderPathExcluded -FolderPath 'Documents\P\CADD\Working\scratch') 'Working excluded'

# oneLevelDeep flag from discovery entry
_Assert (Get-QCAuditSheetsDiscoveryOneLevelDeep -Entry $ht) 'discovery default oneLevelDeep true'
_Assert (-not (Get-QCAuditSheetsDiscoveryOneLevelDeep -Entry @{ FolderPath = 'x'; OneLevelDeep = $false })) 'explicit false'

# watchList.folders config parsing (mirrors reconciliation explicit folders)
$cfgFolders = @{
    projectWise = @{
        watchList = @{
            folders = @(
                @{ root = 'Documents\Root'; path = 'Proj\CADD\Sheets'; oneLevelDeep = $true }
                @{ root = 'Documents\Root'; path = 'Proj\CADD\Sheets'; oneLevelDeep = $false }
            )
        }
    }
}
$entries = @(Get-QCAuditWatchListFolderEntriesFromConfig -Config $cfgFolders)
_Assert ($entries.Count -eq 2) 'two watchList.folders entries'
_Assert ($entries[0].FolderPath -eq 'Documents\Root\Proj\CADD\Sheets') 'joined folder path'

# Path candidates include cmdlet + Documents\ forms (matches Find-PWSheetsFoldersUnderRoot)
$candidates = @(Get-QCAuditCacheWarmFolderPathCandidates -FolderPath 'Documents\AZDOT 2024\Project\CADD\Sheets')
_Assert ($candidates -contains 'AZDOT 2024\Project\CADD\Sheets') 'stripped cmdlet path'
_Assert ($candidates -contains 'Documents\AZDOT 2024\Project\CADD\Sheets') 'Documents\ prefixed path'
$lowerCandidate = @($candidates | Where-Object { $_ -ceq 'documents\azdot 2024\project\cadd\sheets' })
_Assert ($lowerCandidate.Count -eq 0) 'no lowercase API path candidates'

# Child folder GUID in cache -> normal PDF passes parent GUID gate
$cfg = @{ auditPoller = @{ folderGuidCache = @{ filterByParentGuidCache = $true } } }
$childGuid = '9475dfe8-1a85-46de-8986-3e59744591ca'
Register-AuditPollerFolderGuidCacheEntry -Config $cfg -FolderGuid $childGuid -FolderPath $seg
$pdfRow = @{ o_action = 1006; o_parentguid = $childGuid; o_itemname = '015_G-41.01_d0847spp.pdf' }
$gatePdf = Invoke-QCAuditParentGuidCacheGate -Rows @($pdfRow) -Config $cfg -WatchRootConfigs @()
_Assert ($gatePdf.kept.Count -eq 1) 'sheet PDF passes when child folder GUID cached'
_Assert ($gatePdf.diagnostics.passed_parent_cached -eq 1) 'counted as passed_parent_cached'

# 1012 exempt unchanged
$stateRow = @{ o_action = 1012; o_parentguid = '00000000-0000-0000-0000-000000000099'; o_itemname = '0818000063ea515-qc.pdf' }
$gateState = Invoke-QCAuditParentGuidCacheGate -Rows @($stateRow) -Config $cfg
_Assert ($gateState.kept.Count -eq 1) '1012 still exempt from parent GUID filter'

# GUID extraction from [guid]-typed property (pwps_dab sparse folder objects)
$mod = Get-Module PW.AuditPoller
$guidObj = [pscustomobject]@{ FolderPath = 'Documents\Proj\CADD\Sheets'; InstanceGUID = [guid]'9475dfe8-1a85-46de-8986-3e59744591ca' }
$extracted = $mod.Invoke({
    param($Folder)
    _AuditPoller-ExtractFolderGuidFromPwFolder -Folder $Folder
}, @($guidObj))
_Assert ($extracted -eq '9475dfe8-1a85-46de-8986-3e59744591ca') 'extracts [guid] property from folder object'

# Path index reverse lookup (Get-PWFoldersHashTableByGuid fallback)
$indexGuid = '2addfbf1-b111-4afe-8134-8e5f9420d851'
$indexPath = 'documents\azdot 2024\azfwy2302-027_centennial\cadd\sheets'
$mod.Invoke({
    param($Path, $Guid)
    $script:AuditPoller_PwFolderGuidByPath = @{ $Path = $Guid }
}, @($indexPath, $indexGuid)) | Out-Null
$fromIndex = $mod.Invoke({
    param($FolderPath)
    _AuditPoller-TryResolveFolderGuidFromPathIndex -FolderPath $FolderPath
}, @('Documents\AZDOT 2024\AZFWY2302-027_Centennial\CADD\Sheets'))
_Assert ($fromIndex -eq $indexGuid) 'path index resolves normalized folder path to GUID'

Write-Host 'OK: audit folder cache warm tests passed.' -ForegroundColor Green
