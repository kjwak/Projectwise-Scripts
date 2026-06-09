# Unit checks for o_parentguid cache gate (Invoke-QCAuditParentGuidCacheGate).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.AuditPoller.psm1') -Force

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$v = Get-AuditPollerLogicVersion
_Assert ($v -eq '2026-06-05-parent-guid-cache-gate-v10') "Expected parent-guid gate logic version, got: $v"

$watchGuid = '9475dfe8-1a85-46de-8986-3e59744591ca'
$otherGuid = '00000000-0000-0000-0000-000000000099'
$cfg = @{
    auditPoller = @{
        folderGuidCache = @{
            filterByParentGuidCache = $true
        }
    }
}

Register-AuditPollerFolderGuidCacheEntry -Config $cfg -FolderGuid $watchGuid -FolderPath 'Documents\AZDOT 2024\Project\CADD\Sheets'

$rows = @(
    @{ o_action = 1007; o_parentguid = $watchGuid; o_objguid = 'doc-1'; o_itemname = 'a.dgn' },
    @{ o_action = 1007; o_parentguid = $otherGuid; o_objguid = 'doc-2'; o_itemname = 'b.dgn' },
    @{ o_action = 1012; o_parentguid = $otherGuid; o_objguid = 'doc-state'; o_itemname = '0818000063ea515-qc.pdf' },
    @{ o_action = 1002; pw_parentguid = $watchGuid; pw_objguid = 'doc-3'; pw_itemname = 'c.pdf' }
)

$gate = Invoke-QCAuditParentGuidCacheGate -Rows $rows -Config $cfg -WatchRootConfigs @()
_Assert ($gate.active) 'gate should be active when cache has entries'
_Assert ($gate.kept.Count -eq 3) "expected 3 kept rows, got $($gate.kept.Count)"
_Assert ($gate.skipped.Count -eq 1) "expected 1 skipped row, got $($gate.skipped.Count)"

$statsExempt = @{}
$gateState = Invoke-QCAuditParentGuidCacheGate -Rows @($rows[2]) -Config $cfg -Stats $statsExempt
_Assert ($gateState.kept.Count -eq 1) 'DOCUMENT_STATE must pass parent GUID filter even when folder not cached'
_Assert ($statsExempt.parentGuidFilterExemptPassed -eq 1) 'one DOCUMENT_STATE should pass via exempt list'

$disabledCfg = @{
    auditPoller = @{
        folderGuidCache = @{
            filterByParentGuidCache = $false
        }
    }
}
$gateOff = Invoke-QCAuditParentGuidCacheGate -Rows $rows -Config $disabledCfg
_Assert (-not $gateOff.active) 'gate should be inactive when disabled'
_Assert ($gateOff.kept.Count -eq 4) 'disabled gate should pass all rows'

$emptyGate = Invoke-QCAuditParentGuidCacheGate -Rows @() -Config $cfg
_Assert ($emptyGate.kept.Count -eq 0) 'empty input should return empty kept'
_Assert ($emptyGate.skipped.Count -eq 0) 'empty input should return empty skipped'

# Skip-reason classification (row fields only; parent not in cache unless noted)
$uncachedParent = '00000000-0000-0000-0000-000000000099'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = ''; o_itemname = 'a.dgn' }) -eq 'skipped_unknown_parent') 'blank parent'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = $uncachedParent; o_itemname = 'template.itl' }) -eq 'skipped_non_source_extension') 'itl'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = $uncachedParent; o_itemname = '0818000063ea502-qc.pdf' }) -eq 'skipped_qc_artifact') 'dash-qc pdf'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = $uncachedParent; o_itemname = '_StatusSet.pdf' }) -eq 'skipped_status_set_output') 'status set output'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = $uncachedParent; o_itemname = 'e1eb559eef1b9bb9_2265_001_of_001.PwPerfDoNotUse' }) -eq 'skipped_pw_perf') 'PwPerfDoNotUse'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = $uncachedParent; o_itemname = '080J082001ab001.dgn' }) -eq 'skipped_parent_not_cached') 'dgn uncached parent'
_Assert ((Get-QCAuditParentGuidFilterSkipReason -Row @{ o_parentguid = $uncachedParent; o_itemname = '080J082001ab001.pdf' }) -eq 'skipped_parent_not_cached') 'normal pdf uncached parent'

# Gate diagnostics counters and samples (behavior unchanged)
$mixedRows = @(
    @{ o_action = 1007; o_parentguid = $uncachedParent; o_itemname = '080J082001ab001.dgn' },
    @{ o_action = 1007; o_parentguid = $uncachedParent; o_itemname = 'F0893 US 60 Schulze to Allred_ORD Templatel.itl' },
    @{ o_action = 1012; o_parentguid = $uncachedParent; o_itemname = '0818000063ea515-qc.pdf' },
    @{ o_action = 1007; o_parentguid = $watchGuid; o_itemname = '080J082001ab001.pdf' }
)
$mixedStats = @{}
$mixedGate = Invoke-QCAuditParentGuidCacheGate -Rows $mixedRows -Config $cfg -Stats $mixedStats
_Assert ($mixedGate.kept.Count -eq 2) "expected 2 kept (1012 exempt + cached parent), got $($mixedGate.kept.Count)"
_Assert ($mixedGate.skipped.Count -eq 2) "expected 2 skipped, got $($mixedGate.skipped.Count)"
_Assert ($mixedGate.diagnostics.passed_exempt_action -eq 1) 'one exempt 1012'
_Assert ($mixedGate.diagnostics.passed_parent_cached -eq 1) 'one cached parent pass'
_Assert ($mixedGate.diagnostics.skipped_parent_not_cached -eq 1) 'one uncached dgn'
_Assert ($mixedGate.diagnostics.skipped_non_source_extension -eq 1) 'one itl'
_Assert (($mixedGate.diagnostics.skipped_parent_not_cached + $mixedGate.diagnostics.skipped_non_source_extension) -eq $mixedGate.skipped.Count) 'skip counters sum to skipped'
_Assert ($mixedGate.diagnostics.skippedSamples.Count -le 5) 'at most 5 skipped samples'
_Assert ($mixedGate.diagnostics.skippedSamples[0].reason) 'sample includes reason'
_Assert ($mixedStats.passed_exempt_action -eq 1) 'stats mirror passed_exempt_action'
_Assert ($mixedStats.skipped_non_source_extension -eq 1) 'stats mirror skipped_non_source_extension'

$appliedStats = @{}
$appliedGate = Invoke-QCAuditParentGuidCacheGate -Rows $rows -Config $cfg -Stats $appliedStats
_Assert ($appliedGate.active) 'active gate should set filter_applied telemetry'
_Assert ($appliedStats.filterByParentGuidCacheConfigured) 'configured true when filter enabled'
_Assert ($appliedStats.folderGuidCacheConfigPresent) 'folderGuidCache block present'
_Assert ($appliedStats.parentGuidFilterActivationReason -eq 'filter_applied') 'activation reason when gate runs'
_Assert ($null -eq $appliedStats.parentGuidFilterBypassReason) 'no bypass when gate runs'

$disabledStats = @{}
$disabledGate = Invoke-QCAuditParentGuidCacheGate -Rows $rows -Config $disabledCfg -Stats $disabledStats
_Assert (-not $disabledGate.active) 'disabled gate inactive'
_Assert (-not $disabledStats.filterByParentGuidCacheConfigured) 'configured false when filter disabled'
_Assert ($disabledStats.parentGuidFilterBypassReason -eq 'disabled') 'disabled bypass reason'
_Assert ($null -eq $disabledStats.parentGuidFilterActivationReason) 'no activation when disabled'

Write-Host 'OK: audit parent GUID filter tests passed.' -ForegroundColor Green
