$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.AuditPoller.psm1') -Force

$cfgRes = Read-QCAppSettings -Path (Join-Path $repoRoot 'appsettings.json')
$config = $cfgRes.Data.config
$pwCfg = ConvertTo-HashtableDeep -Value $config.projectWise
$watchList = ConvertTo-HashtableDeep -Value $pwCfg.watchList
$rootsRaw = $watchList.roots
Write-Host "roots type: $($rootsRaw.GetType().FullName)"
Write-Host "roots count: $(@($rootsRaw).Count)"

$watchRootConfigs = @($watchList.roots | ForEach-Object { ConvertTo-HashtableDeep -Value $_ })
Write-Host "watchRootConfigs count: $($watchRootConfigs.Count)"

$norm = @(_AuditPoller-NormalizeWatchRootConfigs -WatchRootConfigs $watchRootConfigs -Config $config)
Write-Host "normalized count: $($norm.Count)"
$watchRoots = @($norm | ForEach-Object { _AuditPoller-GetWatchRootPathFromConfig -Cfg $_ })
foreach ($r in $watchRoots) { Write-Host "  root: $r" }

$matchRoots = _AuditPoller-BuildMatchRoots -WatchRoots $watchRoots
$fp = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
$match = _AuditPoller-MatchesWatchRoot -FolderPath $fp -MatchRoots $matchRoots
Write-Host "Seg_1 match: $match"

# Param binding pitfall: pass array of one root without wrapping
$paramTest = $watchRootConfigs[0]
Write-Host "paramTest type: $($paramTest.GetType().FullName)"
function Test-ArrayParam { param([array]$WatchRootConfigs) 
    Write-Host "  param count: $($WatchRootConfigs.Count)"
    Write-Host "  param[0] type: $($WatchRootConfigs[0].GetType().FullName)"
}
Test-ArrayParam -WatchRootConfigs $paramTest
Test-ArrayParam -WatchRootConfigs @($paramTest)
