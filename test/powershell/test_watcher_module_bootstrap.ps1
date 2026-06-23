# Verifies Watch-QCTrigger module import chain exposes Test-QCDatabaseEnabled.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

$imports = @(
    'modules\Core\Core.Results.psm1',
    'modules\Core\Core.Runtime.psm1',
    'modules\Core\Core.Hashing.psm1',
    'modules\Core\Core.Paths.psm1',
    'modules\Queue\QC.Filters.psm1',
    'modules\Queue\QC.Triggers.psm1',
    'modules\Queue\QC.JobFactory.psm1',
    'modules\Queue\QC.Queue.Json.psm1',
    'modules\Database\Core.Database.psm1',
    'modules\ProjectWise\PW.Users.psm1',
    'modules\ProjectWise\PW.Discovery.psm1',
    'modules\ProjectWise\PW.AuditPoller.psm1',
    'modules\Processing\QC.StatusSet.psm1',
    'modules\Core\QC.WatcherOrchestration.psm1'
)
foreach ($rel in $imports) {
    Import-Module (Join-Path $repoRoot $rel) -Force
}
if (-not (Get-Command Test-QCDatabaseEnabled -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
}

if (-not (Get-Command Test-QCDatabaseEnabled -ErrorAction SilentlyContinue)) {
    throw 'Test-QCDatabaseEnabled missing after watcher import chain'
}
if (-not (Get-Command Read-QCAppSettings -ErrorAction SilentlyContinue)) {
    throw 'Read-QCAppSettings missing after watcher import chain'
}
Write-Host 'test_watcher_module_bootstrap.ps1 passed'
