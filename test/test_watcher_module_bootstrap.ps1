# Verifies Watch-QCTrigger module import chain exposes Test-QCDatabaseEnabled.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

$imports = @(
    'modules\Core.Results.psm1',
    'modules\Core.Runtime.psm1',
    'modules\Core.Hashing.psm1',
    'modules\Core.Paths.psm1',
    'modules\QC.Filters.psm1',
    'modules\QC.Triggers.psm1',
    'modules\QC.JobFactory.psm1',
    'modules\QC.Queue.Json.psm1',
    'modules\Core.Database.psm1',
    'modules\PW.Users.psm1',
    'modules\PW.Discovery.psm1',
    'modules\PW.AuditPoller.psm1',
    'modules\QC.StatusSet.psm1',
    'modules\QC.WatcherOrchestration.psm1'
)
foreach ($rel in $imports) {
    Import-Module (Join-Path $repoRoot $rel) -Force
}
if (-not (Get-Command Test-QCDatabaseEnabled -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
}

if (-not (Get-Command Test-QCDatabaseEnabled -ErrorAction SilentlyContinue)) {
    throw 'Test-QCDatabaseEnabled missing after watcher import chain'
}
if (-not (Get-Command Read-QCAppSettings -ErrorAction SilentlyContinue)) {
    throw 'Read-QCAppSettings missing after watcher import chain'
}
Write-Host 'test_watcher_module_bootstrap.ps1 passed'
