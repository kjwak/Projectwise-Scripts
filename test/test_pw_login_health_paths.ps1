$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/PW.Connection.psm1" -Force

InModuleScope -ModuleName PW.Connection {
    Assert-Eq (_PWC-ConvertProbeFolderPath 'Documents\AZDOT 2024') 'AZDOT 2024' 'watch root path should strip Documents\ prefix'
    Assert-Eq (_PWC-ConvertProbeFolderPath 'Documents') 'Documents' 'Documents root should stay Documents'

    $config = @{
        projectWise = @{
            watchList = @{
                roots = @(
                    @{ path = 'Documents\AZDOT 2024' }
                    @{ path = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE' }
                )
            }
        }
    }

    $candidates = _PWC-GetSessionProbeFolderCandidates -Config $config
    Assert-Eq $candidates[0] 'Documents' 'Documents should be first probe candidate'
    Assert-True ($candidates -contains 'AZDOT 2024') 'converted watch root should be included'
    Assert-True ($candidates -contains 'Caltrans\CAFWY2200-I-15_ELPSE') 'second watch root should be included'
}

Write-Host 'OK: PW login health path tests passed.' -ForegroundColor Green
