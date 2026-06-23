# Unit tests for QC_COMMENT_STATUS_SYNC trigger rule matching (committed appsettings lane model).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules/Core/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Core/Core.Paths.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Queue/QC.Triggers.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Core/Core.Runtime.psm1') -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-MatchRule($Rule, $FileName, $ShouldMatch, $Label) {
    $candidate = @{
        path     = "Documents\Proj\CADD\Sheets\$FileName"
        fileName = $FileName
        description = ''
    }
    $r = Test-TriggerRule -Candidate $candidate -Rule $Rule
    Assert-True $r.IsSuccess "Trigger test should succeed for $Label"
    $isMatch = [bool]$r.Data.isMatch
    if ($ShouldMatch) {
        Assert-True $isMatch "$Label should match QC_COMMENT_STATUS_SYNC rule"
    } else {
        Assert-True (-not $isMatch) "$Label should not match QC_COMMENT_STATUS_SYNC rule"
    }
}

$cfgRes = Read-QCAppSettings -Path (Join-Path $repoRoot 'appsettings.json')
Assert-True $cfgRes.IsSuccess 'appsettings.json should load'
$config = [hashtable]$cfgRes.Data.config

$rule = $null
foreach ($r in @($config.triggers.rules)) {
    if ([string]$r.jobType -eq 'QC_COMMENT_STATUS_SYNC') {
        $rule = $r
        break
    }
}
Assert-True ($null -ne $rule) 'Committed appsettings should define QC_COMMENT_STATUS_SYNC trigger rule'

Assert-MatchRule $rule 'sheet001-prod.pdf' $true 'lane prod'
Assert-MatchRule $rule 'sheet001-rev.pdf'  $true 'lane rev'
Assert-MatchRule $rule 'sheet001-chk.pdf'  $true 'lane chk'
Assert-MatchRule $rule 'sheet001.pdf'      $false 'stem sheet PDF'
Assert-MatchRule $rule 'sheet001-qc.pdf'   $false 'legacy *-qc.pdf (not primary trigger)'

# Legacy regex would match -qc.pdf; production rule must not.
$legacyRule = @{
    id = 'qc-comment-status-legacy'
    jobType = 'QC_COMMENT_STATUS_SYNC'
    triggerType = 'pw'
    when = @{
        extensions = @('.pdf')
        fileNameRegexAny = @('(?i)-qc\.pdf$')
    }
    requireAll = @('extensions', 'fileNameRegexAny')
    exclude = @{ pathRegexAny = @(); fileNameRegexAny = @() }
}
Assert-MatchRule $legacyRule 'sheet001-qc.pdf' $true 'legacy rule still matches -qc.pdf when configured'

Write-Host 'OK test_qc_comment_trigger.ps1'
