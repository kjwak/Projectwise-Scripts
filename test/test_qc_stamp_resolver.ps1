$ErrorActionPreference = 'Stop'
function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force

$config = @{
    QCProcess = @{
        DefaultProcessType = 'production'
        ProcessTypes = @{
            production = @{ PdfSuffix = 'prod'; DefaultStamp = 'Production' }
            check = @{ PdfSuffix = 'chk'; DefaultStamp = 'Check' }
            review = @{ PdfSuffix = 'rev'; DefaultStamp = 'Review' }
        }
        StampProfiles = @{
            Default = @{ production = 'Production'; check = 'Check'; review = 'Review' }
            ProjectA = @{ production = 'Production'; check = 'Check'; review = 'Review' }
        }
        StampAssets = @{
            Production = 'stamps/Production_Stamp.pdf'
            Check = 'stamps/IC_Stamp.pdf'
            Review = 'stamps/Peer_Review_Stamp.pdf'
        }
        RootOverrides = @(
            @{ Name = 'Broad'; RootPath = 'Documents/Projects'; StampProfile = 'Default' }
            @{ Name = 'Nested'; RootPath = 'Documents/Projects/Example/Nested'; StampProfile = 'ProjectA' }
        )
    }
}

$default = Resolve-QCStampForProcess -Config $config -ProcessType 'check' -FolderPath 'Documents/Other'
Assert-True $default.IsSuccess 'default check stamp resolves'

$nested = Resolve-QCStampForProcess -Config $config -ProcessType 'review' -FolderPath 'Documents/Projects/Example/Nested/Sheets'
Assert-Eq $nested.resolvedStampProfile 'ProjectA' 'longest-prefix override'

$missing = Resolve-QCStampForProcess -Config @{
    QCProcess = @{
        ProcessTypes = @{ check = @{ DefaultStamp = 'MissingStamp' } }
        StampProfiles = @{ Default = @{} }
        StampAssets = @{}
    }
    qcPrepend = @{ reviewStamps = @{ enabled = $false } }
} -ProcessType 'check'
Assert-True (-not $missing.IsSuccess) 'missing stamp profile fails'

Write-Host 'test_qc_stamp_resolver.ps1: OK'
