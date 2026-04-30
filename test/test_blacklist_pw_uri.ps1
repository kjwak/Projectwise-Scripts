<#
.SYNOPSIS
Regression test: a blacklist entry written as a "pw:\\datasource\Documents\..."
URI must successfully filter a candidate path produced by the watcher in plain
"Documents\..." form. The previous Normalize-QCPath logic stripped the "pw:" prefix
only when the path still contained "\\\\" (literal "\\") after a "\\{2,}" collapse,
which silently failed for JSON-encoded entries like "pw:\\\\datasource\\Documents\\...".
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Core.Paths.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\QC.Filters.psm1') -Force -DisableNameChecking | Out-Null

function Assert-True($Cond, $Msg) {
    if (-not $Cond) { throw "ASSERT FAILED: $Msg" }
}

# Mirrors the real appsettings.json shape after JSON parse:
#   "pw:\\\\typsa...\\Documents\\AZDOT\\AZFWY2211_I17-303L-TI\\"
# decodes to:
#   pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\AZDOT\AZFWY2211_I17-303L-TI\
$blacklistRoot = "pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\AZDOT\AZFWY2211_I17-303L-TI\"

# 1) Normalization must reduce the URI to "Documents\AZDOT\AZFWY2211_I17-303L-TI"
$norm = Normalize-QCPath -Path $blacklistRoot
Assert-True $norm.IsSuccess 'Normalize-QCPath should succeed'
Assert-True ([string]$norm.Data.path -eq 'Documents\AZDOT\AZFWY2211_I17-303L-TI') ("Normalized PW URI mismatch: '$($norm.Data.path)'")

# 2) A candidate Sheets path under that project must register as under-root.
$candidate = 'Documents\AZDOT\AZFWY2211_I17-303L-TI\CADD\Sheets'
$under = Test-PathUnderRoot -Path $candidate -Root $blacklistRoot
Assert-True ($under.IsSuccess -and [bool]$under.Data.isUnderRoot) "Candidate '$candidate' should be under blacklist root '$blacklistRoot'"

# 3) End-to-end: Test-QCPathAllowed must block the candidate.
$config = @{
    filters = @{
        whitelist = @{ enabled = $false; paths = @() }
        blacklist = @{
            paths = @($blacklistRoot)
            patterns = @()
        }
    }
}
$res = Test-QCPathAllowed -CandidatePath $candidate -Config $config
Assert-True $res.IsSuccess 'Test-QCPathAllowed should succeed'
Assert-True (-not [bool]$res.Data.allowed) "Candidate '$candidate' should be FILTERED but was ALLOWED (reason='$($res.Data.reason)')"
Assert-True ($res.Code -eq 'FILTERED_BLACKLIST') "Expected code FILTERED_BLACKLIST, got '$($res.Code)'"

# 4) A candidate NOT under the blacklisted project must still be allowed.
$other = 'Documents\AZDOT\AZFWY2302-005\CADD\Sheets'
$resOk = Test-QCPathAllowed -CandidatePath $other -Config $config
Assert-True ($resOk.IsSuccess -and [bool]$resOk.Data.allowed) "Unrelated path '$other' should be allowed"

Write-Host "test_blacklist_pw_uri: PASS" -ForegroundColor Green
