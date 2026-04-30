<#
.SYNOPSIS
Regression: Invoke-StatusSetProcessor must NOT silently succeed when statusSet
config is missing or contains an unrecognized mode. The previous default
('stub') caused jobs to "succeed" in <1s without any work.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\QC.Processors.psm1') -Force -DisableNameChecking | Out-Null

function Assert-True($Cond, $Msg) {
    if (-not $Cond) { throw "ASSERT FAILED: $Msg" }
}

$job = @{
    id           = 'test_ssproc_default'
    type         = 'STATUS_SET_GEN'
    sourceFolder = 'Documents\AZDOT\TESTPROJ\CADD\Sheets'
    metadata     = @{ candidate = @{ datasourceName = 'unit-test'; oneLevelDeep = $true } }
}

# 1) No statusSet block: must default to 'native' (not 'stub'). Native dispatch will
# attempt to load QC.StatusSet.psm1 and call PW; it WILL fail downstream because we
# have no PW connection, but the failure must be from the native path, not a silent
# stub success. Accept any non-stub-success outcome.
$cfg1 = @{}
$res1 = Invoke-StatusSetProcessor -Job $job -Config $cfg1
Assert-True ($res1 -ne $null) 'Result must not be null'
Assert-True ($res1.Code -ne 'STATUS_SET_STUB_OK') ("Default mode should NOT silently stub-succeed; got Code='{0}'" -f $res1.Code)
Assert-True (($res1.Code -like 'STATUS_SET_*') -or ($res1.Code -like 'PW_*')) ("Expected a STATUS_SET_* or PW_* code, got '{0}'" -f $res1.Code)

# 2) Unknown mode: must FAIL loudly, not succeed.
$cfg2 = @{ statusSet = @{ mode = 'banana' } }
$res2 = Invoke-StatusSetProcessor -Job $job -Config $cfg2
Assert-True ($res2 -ne $null) 'Result must not be null'
Assert-True (-not $res2.IsSuccess) 'Unknown statusSet.mode must NOT succeed'
Assert-True ($res2.Code -eq 'STATUS_SET_UNKNOWN_MODE') ("Expected STATUS_SET_UNKNOWN_MODE, got '{0}'" -f $res2.Code)

# 3) Explicit mode='stub' is allowed (intentional opt-out) and remains a no-op success.
$cfg3 = @{ statusSet = @{ mode = 'stub' } }
$res3 = Invoke-StatusSetProcessor -Job $job -Config $cfg3
Assert-True ($res3 -ne $null) 'Result must not be null'
Assert-True ($res3.IsSuccess) 'Explicit stub mode should succeed (intentional opt-out)'
Assert-True ($res3.Code -eq 'STATUS_SET_STUB_OK') ("Expected STATUS_SET_STUB_OK, got '{0}'" -f $res3.Code)

Write-Host "test_statusset_processor_default: PASS" -ForegroundColor Green
