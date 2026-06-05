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
_Assert ($v -eq '2026-06-05-parent-guid-cache-gate-v9') "Expected parent-guid gate logic version, got: $v"

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
    @{ o_action = 1012; o_parentguid = $otherGuid; o_objguid = 'doc-2'; o_itemname = 'b.pdf' },
    @{ o_action = 1002; pw_parentguid = $watchGuid; pw_objguid = 'doc-3'; pw_itemname = 'c.pdf' }
)

$gate = Invoke-QCAuditParentGuidCacheGate -Rows $rows -Config $cfg -WatchRootConfigs @()
_Assert ($gate.active) 'gate should be active when cache has entries'
_Assert ($gate.kept.Count -eq 2) "expected 2 kept rows, got $($gate.kept.Count)"
_Assert ($gate.skipped.Count -eq 1) "expected 1 skipped row, got $($gate.skipped.Count)"

$disabledCfg = @{
    auditPoller = @{
        folderGuidCache = @{
            filterByParentGuidCache = $false
        }
    }
}
$gateOff = Invoke-QCAuditParentGuidCacheGate -Rows $rows -Config $disabledCfg
_Assert (-not $gateOff.active) 'gate should be inactive when disabled'
_Assert ($gateOff.kept.Count -eq 3) 'disabled gate should pass all rows'

Write-Host 'OK: audit parent GUID filter tests passed.' -ForegroundColor Green
