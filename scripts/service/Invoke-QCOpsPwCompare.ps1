<#
.SYNOPSIS
Child helper: ProjectWise vs SQL compare in an MTA process. Writes JSON to -OutputPath.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AppSettingsPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$SheetNumber = '',
    [string]$DocumentGuid = ''
)

$ErrorActionPreference = 'Stop'
$scriptsRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $scriptsRoot

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -RepoRoot $repoRoot -FeatureModules @(
    'Core\Core.Results.psm1'
    'Database\Core.Database.psm1'
    'Diagnostics\QC.DebugMcp.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Initialize-QCDebugMcpContext'
    'Compare-QCProjectWiseToDatabase'
) -Context 'ops PW compare child'

Initialize-QCDebugMcpContext -AppSettingsPath $AppSettingsPath | Out-Null
$result = $null
if ($DocumentGuid) {
    $result = Compare-QCProjectWiseToDatabase -DocumentGuid $DocumentGuid
} else {
    $result = Compare-QCProjectWiseToDatabase -SheetNumber $SheetNumber
}
$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
($result | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $OutputPath -Encoding utf8
