<#
.SYNOPSIS
Opens a log console at user logon (Startup folder shortcut).

.DESCRIPTION
Creates a Startup shortcut that launches Start-QCOpsConsole.ps1 (WinForms ops GUI)
in a visible PowerShell window. Does not start the dashboard (that is the
QC-PipelineDashboard scheduled task). Pass -NoGui on the console script for the
text log tail.

.EXAMPLE
.\scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1

.EXAMPLE
.\scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1 -Unregister
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$ShortcutName = 'QC Pipeline Dashboard Logs.lnk',

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = '',

    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = '',

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

$expectedHost = 'PXBENTLEY01'
$modellingHost = 'AZTEC002799'
$prodCloneRoot = 'C:\Users\jflint\Documents\github\Prepend PDF QC'

function _Quote-TaskArg {
    param([Parameter(Mandatory)][string]$Value)
    if ($Value -notmatch '[ \t"]') { return $Value }
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

$scriptsRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $hostName = [string]$env:COMPUTERNAME
    $prodDash = Join-Path $prodCloneRoot 'scripts\service\Start-QCPipelineDashboard.ps1'
    if ($hostName -ieq $expectedHost -and (Test-Path -LiteralPath $prodDash)) {
        $RepoRoot = $prodCloneRoot
    } else {
        $RepoRoot = Split-Path -Parent $scriptsRoot
    }
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$guiScript = Join-Path $RepoRoot 'scripts\service\Start-QCOpsConsole.ps1'
$watchScript = Join-Path $RepoRoot 'scripts\service\Watch-QCPipelineDashboardConsole.ps1'
$launchScript = $guiScript
if (-not (Test-Path -LiteralPath $guiScript)) {
    $launchScript = $watchScript
}
if (-not (Test-Path -LiteralPath $launchScript)) {
    throw "Ops console script not found: $guiScript"
}

$hostName = [string]$env:COMPUTERNAME
if (-not $Force.IsPresent -and -not $Unregister.IsPresent) {
    if ($hostName -ieq $modellingHost) {
        throw "Refusing to register the QC server log console on modelling host $hostName. Use Register-QCRemoteWorkerHostLogonConsole.ps1 there, or pass -Force to override."
    }
    if ($hostName -ine $expectedHost) {
        throw "This shortcut is for QC server $expectedHost (this host is $hostName). RDP to $expectedHost and re-run, or pass -Force to override."
    }
}

$startupDir = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupDir $ShortcutName

if ($Unregister.IsPresent) {
    if ($PSCmdlet.ShouldProcess($shortcutPath, 'Remove Startup shortcut')) {
        Remove-Item -LiteralPath $shortcutPath -Force -ErrorAction SilentlyContinue
        Write-Host "Removed Startup shortcut: $shortcutPath" -ForegroundColor Green
    }
    return
}

if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $RepoRoot 'appsettings.json'
}
$AppSettingsPath = (Resolve-Path -LiteralPath $AppSettingsPath).Path

$argList = New-Object System.Collections.Generic.List[string]
[void]$argList.Add('-STA')
[void]$argList.Add('-NoProfile')
[void]$argList.Add('-ExecutionPolicy')
[void]$argList.Add('Bypass')
[void]$argList.Add('-File')
[void]$argList.Add((_Quote-TaskArg -Value $launchScript))
[void]$argList.Add('-AppSettingsPath')
[void]$argList.Add((_Quote-TaskArg -Value $AppSettingsPath))

if ($PSCmdlet.ShouldProcess($shortcutPath, 'Create Startup shortcut')) {
    $wsh = New-Object -ComObject WScript.Shell
    $sc = $wsh.CreateShortcut($shortcutPath)
    $sc.TargetPath = 'powershell.exe'
    $sc.Arguments = ($argList -join ' ')
    $sc.WorkingDirectory = $RepoRoot
    $sc.WindowStyle = 1
    $sc.Description = 'QC pipeline ops console (read-only + task on/off). Does not start a second dashboard.'
    $sc.Save()

    Write-Host "Created Startup shortcut: $shortcutPath" -ForegroundColor Green
    Write-Host "  Opens at logon: $launchScript" -ForegroundColor Gray
    Write-Host '  Does not start the dashboard (QC-PipelineDashboard scheduled task does that).' -ForegroundColor Gray
    Write-Host '  Text tail fallback: Start-QCOpsConsole.ps1 -NoGui' -ForegroundColor Gray
    Write-Host '  Remove: .\scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1 -Unregister' -ForegroundColor Cyan
}
