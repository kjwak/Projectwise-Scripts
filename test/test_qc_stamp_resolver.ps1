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
            I15_ELPSE = @{
                production = 'Production'
                check = 'Check'
                review = 'Review'
                layout = @{
                    stampHeightPt = 250
                    stampPositionPt = @{ x = -350; y = 10 }
                }
            }
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

$layoutDefault = Get-QCStampProfileLayout -Config @{
    qcPrepend = @{ reviewStamps = @{ stampHeightPt = 300; stampPositionPt = @{ x = -400; y = 0 } } }
    QCProcess = @{ StampProfiles = @{ Default = @{} } }
} -StampProfile 'Default'
Assert-Eq $layoutDefault.stampHeightPt 300 'default layout inherits global height'
Assert-Eq $layoutDefault.stampXPt -400 'default layout inherits global x'

$layoutI15 = Get-QCStampProfileLayout -Config $config -StampProfile 'I15_ELPSE'
Assert-Eq $layoutI15.stampHeightPt 250 'profile layout overrides height'
Assert-Eq $layoutI15.stampXPt -350 'profile layout overrides x'

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
