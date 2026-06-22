<# Launcher: loads appsettings.json defaults and runs Combine-StatusSet.ps1 #>
param(
    [Parameter(Mandatory = $true)]
    [string] $SheetsFolderPath,

    [Parameter(Mandatory = $false)]
    [string] $AppSettingsPath = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'appsettings.json'),

    [Parameter(Mandatory = $false)]
    [bool] $OneLevelDeep = $true,

    [Parameter(Mandatory = $false)]
    [switch] $ForceRebuild,

    [Parameter(Mandatory = $false)]
    [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

$localRoot = 'C:\PW_QC_LOCAL'
if (Test-Path -LiteralPath $AppSettingsPath) {
    $raw = Get-Content -LiteralPath $AppSettingsPath -Raw | ConvertFrom-Json
    if ($raw.statusSet -and $raw.statusSet.localRoot) { $localRoot = [string]$raw.statusSet.localRoot }
}

$target = Join-Path $PSScriptRoot 'Combine-StatusSet.ps1'
& $target -SheetsFolderPath $SheetsFolderPath -LocalRoot $localRoot -OneLevelDeep:$OneLevelDeep -ForceRebuild:$ForceRebuild -DryRun:$DryRun
exit $LASTEXITCODE
