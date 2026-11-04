# Verify pw-qc-debug MCP worker module bootstrap after Phase 4H shim removal.
$ErrorActionPreference = 'Stop'

$repoRoot = if ($env:PWQC_REPO_ROOT) { $env:PWQC_REPO_ROOT } else { Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$scriptsRoot = Join-Path $repoRoot 'scripts'

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Diagnostics\QC.DebugMcp.psm1'
    'Core\Core.Telemetry.psm1'
) -RequiredCommands @(
    'Initialize-QCDebugMcpContext'
    'Get-QCAppSettingsConfig'
    'Search-QCDebugSheet'
    'Get-QCDebugRecentErrors'
    'Get-QCDebugProcessHealth'
    'Compare-QCProjectWiseToDatabase'
) -Context 'pw-qc-mcp bootstrap test'

Import-QCModuleGlobal -RelativePath 'ProjectWise\PW.Connection.psm1'
if (-not (Get-Command -Name 'Invoke-PWAuthenticatedCommand' -ErrorAction SilentlyContinue)) {
    throw 'Invoke-PWAuthenticatedCommand missing after PW.Connection import'
}

$flatShims = @(
    (Join-Path $repoRoot 'modules\QC.DebugMcp.psm1')
    (Join-Path $repoRoot 'modules\Core.Runtime.psm1')
)
foreach ($path in $flatShims) {
    if (Test-Path -LiteralPath $path) {
        throw "Flat module shim should be removed: $path"
    }
}

Write-Host 'test_mcp_bootstrap.ps1 passed' -ForegroundColor Green
exit 0
