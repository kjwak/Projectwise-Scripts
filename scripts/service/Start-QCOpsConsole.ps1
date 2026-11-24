<#
.SYNOPSIS
Logon-session WinForms operator console for the QC pipeline.
#>
[CmdletBinding()]
param([string]$AppSettingsPath, [switch]$Force, [switch]$NoGui, [switch]$SmokeTest)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $AppSettingsPath) { $AppSettingsPath = Join-Path $repoRoot 'appsettings.json' }
$AppSettingsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AppSettingsPath)
if ($NoGui) {
    & (Join-Path $PSScriptRoot 'Watch-QCPipelineDashboardConsole.ps1') -AppSettingsPath $AppSettingsPath
    return
}

$opsModules = Join-Path $repoRoot 'modules\Ops'
Import-Module (Join-Path $opsModules 'QC.OpsProcess.psm1') -DisableNameChecking
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $arguments = @('-STA','-NoProfile','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-AppSettingsPath',$AppSettingsPath)
    if ($Force) { $arguments += '-Force' }
    if ($SmokeTest) { $arguments += '-SmokeTest' }
    $arguments = ($arguments | ForEach-Object { ConvertTo-QCOpsWindowsArgument $_ }) -join ' '
    Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments
    return
}

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -DisableNameChecking
Import-Module (Join-Path $opsModules 'QC.OpsConsole.psm1') -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if (-not (Test-QCOpsHostAllowed -Force:$Force)) {
    [void][Windows.Forms.MessageBox]::Show("QC Ops runs on $(Get-QCOpsExpectedHost).", 'QC Pipeline Ops')
    exit 2
}
[Windows.Forms.Application]::EnableVisualStyles()
$script:Cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath
$script:Cfg['_appSettingsPath'] = $AppSettingsPath
. (Join-Path $PSScriptRoot 'QCOpsConsole.Gui.ps1')
