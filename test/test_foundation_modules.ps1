# Lightweight test script for pure foundation modules.
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/Core.Results.psm1" -Force
Import-Module "$PSScriptRoot/../modules/Core.Paths.psm1" -Force
Import-Module "$PSScriptRoot/../modules/QC.Filters.psm1" -Force
Import-Module "$PSScriptRoot/../modules/QC.Triggers.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

# result contract shape
$rr = New-QCSuccessResult -Code 'OK' -Message 'ok' -Data @{x=1}
Assert-True (($rr.PSObject.Properties.Name -contains 'IsSuccess') -and ($rr.PSObject.Properties.Name -contains 'Code') -and ($rr.PSObject.Properties.Name -contains 'Message') -and ($rr.PSObject.Properties.Name -contains 'Data')) 'Result contract keys should exist'

# path normalization (mixed slashes/case/trailing)
$r = Normalize-QCPath -Path ' Documents//AZDOT 2024\\ProjA\\ '
Assert-True $r.IsSuccess 'Normalize-QCPath should succeed'
Assert-True ($r.Data.path -eq 'documents\azdot 2024\proja') 'Normalize-QCPath should normalize separators/case/trim'

# ProjectWise-style path normalization
$pw = Normalize-QCPath -Path 'pw:/Server:Ds/Documents/Folder/Sub/'
Assert-True $pw.IsSuccess 'ProjectWise-style path should normalize'
Assert-True ($pw.Data.path -eq 'pw:\server:ds\documents\folder\sub') 'ProjectWise-style path normalization mismatch'

# Test-PathUnderRoot should be case-insensitive and slash-agnostic
$ur = Test-PathUnderRoot -Path 'Documents/AZDOT 2024/ProjA/CADD' -Root 'documents\azdot 2024\proja'
Assert-True ($ur.IsSuccess -and $ur.Data.isUnderRoot) 'Test-PathUnderRoot should be true for case/slash variants'

# whitelist allowed
$config = @{
    filters = @{
        whitelist = @{ enabled = $true; paths = @('documents\azdot 2024\proja') }
        blacklist = @{ paths = @(); patterns = @() }
    }
}
$f1 = Test-QCPathAllowed -CandidatePath 'Documents\AZDOT 2024\ProjA\CADD\Sheets' -Config $config
Assert-True ($f1.IsSuccess -and $f1.Data.allowed) 'Whitelist path should be allowed'

# whitelist present means block unless matched
$f1b = Test-QCPathAllowed -CandidatePath 'Documents\AZDOT 2024\ProjB\CADD\Sheets' -Config $config
Assert-True ($f1b.IsSuccess -and -not $f1b.Data.allowed -and $f1b.Data.reason -eq 'not_whitelisted') 'Non-whitelisted path should be blocked when whitelist enabled'

# blacklist blocked when no whitelist
$config.filters.whitelist.enabled = $false
$config.filters.blacklist.paths = @('documents\azdot 2024\proja\cadd\sheets\archive')
$f2 = Test-QCPathAllowed -CandidatePath 'Documents\AZDOT 2024\ProjA\CADD\Sheets\Archive\A.pdf' -Config $config
Assert-True ($f2.IsSuccess -and -not $f2.Data.allowed) 'Blacklist path should be blocked'

# blacklist overrides whitelist
$config.filters.whitelist.enabled = $true
$config.filters.whitelist.paths = @('documents\azdot 2024\proja')
$f3 = Test-QCPathAllowed -CandidatePath 'Documents\AZDOT 2024\ProjA\CADD\Sheets\Archive\A.pdf' -Config $config
Assert-True ($f3.IsSuccess -and -not $f3.Data.allowed -and $f3.Code -eq 'FILTERED_BLACKLIST_OVERRIDE') 'Blacklist should override whitelist'

# trigger rules: disabled ignored / priority / exclude / requireAll / no match
$triggerConfig = @{
    triggers = @{
        rules = @(
            @{ id='disabled-top'; enabled=$false; priority=999; jobType='DISABLED'; triggerType='desc'; when=@{ extensions=@('.pdf'); descriptionContainsAny=@('|QC|'); pathRegexAny=@(); fileNameRegexAny=@() }; requireAll=@('extensions','descriptionContainsAny'); exclude=@{ pathRegexAny=@(); fileNameRegexAny=@() } },
            @{ id='low'; enabled=$true; priority=10; jobType='LOW'; triggerType='desc'; when=@{ extensions=@('.pdf'); descriptionContainsAny=@('|QC|'); pathRegexAny=@(); fileNameRegexAny=@() }; requireAll=@('extensions','descriptionContainsAny'); exclude=@{ pathRegexAny=@(); fileNameRegexAny=@() } },
            @{ id='high'; enabled=$true; priority=100; jobType='HIGH'; triggerType='desc'; when=@{ extensions=@('.pdf'); descriptionContainsAny=@('|QC|'); pathRegexAny=@(); fileNameRegexAny=@() }; requireAll=@('extensions','descriptionContainsAny'); exclude=@{ pathRegexAny=@(); fileNameRegexAny=@() } },
            # "Exclude" rule should only activate for drafts; when it activates, it excludes and lets next match win.
            @{ id='exclude'; enabled=$true; priority=200; jobType='X'; triggerType='name'; when=@{ extensions=@('.pdf'); descriptionContainsAny=@('|QC|'); pathRegexAny=@(); fileNameRegexAny=@('(?i)draft') }; requireAll=@('extensions','descriptionContainsAny','fileNameRegexAny'); exclude=@{ pathRegexAny=@(); fileNameRegexAny=@('(?i)draft') } }
        )
    }
}

$candidate = @{ path='Documents\AZDOT 2024\ProjA\CADD\Sheets\A101.pdf'; fileName='A101.pdf'; description='|QC| ready' }
$match = Test-QCTriggerCandidate -Candidate $candidate -Config $triggerConfig
Assert-True ($match.IsSuccess -and $match.Data.matched) 'Candidate should match a trigger rule'
Assert-True ($match.Data.evaluation.ruleId -eq 'high') 'Disabled rule must be ignored and higher-priority enabled match should win'

$candidateEx = @{ path='Documents\AZDOT 2024\ProjA\CADD\Sheets\Draft.pdf'; fileName='Draft.pdf'; description='|QC| ready' }
$matchEx = Test-QCTriggerCandidate -Candidate $candidateEx -Config $triggerConfig
Assert-True ($matchEx.IsSuccess -and $matchEx.Data.matched -and $matchEx.Data.evaluation.ruleId -eq 'high') 'Excluded higher-priority rule should be skipped and next match should win'

# requireAll failure (description missing marker)
$candidateRequireFail = @{ path='Documents\AZDOT 2024\ProjA\CADD\Sheets\A101.pdf'; fileName='A101.pdf'; description='ready' }
$requireFail = Test-QCTriggerCandidate -Candidate $candidateRequireFail -Config $triggerConfig
Assert-True ($requireFail.IsSuccess -and -not $requireFail.Data.matched) 'requireAll should fail when required description marker missing'

# no match returns clean ignored result
$candidateNo = @{ path='Documents\AZDOT 2024\ProjA\CADD\Sheets\A101.dgn'; fileName='A101.dgn'; description='none' }
$no = Resolve-QCTriggerMatch -Candidate $candidateNo -Config $triggerConfig
Assert-True ($no.IsSuccess -and $no.Code -eq 'IGNORED_NO_MATCH' -and $no.Data.action -eq 'ignore') 'No match should return ignored/no_match'

Write-Host 'All foundation module tests passed.' -ForegroundColor Green
