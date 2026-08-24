<#
.SYNOPSIS
Opens a log console at user logon (Startup folder shortcut).

.DESCRIPTION
Creates a Startup shortcut that launches Watch-QCRemoteWorkerHostConsole.ps1 in a
visible PowerShell window. Does not start the worker supervisor (that is the
QC-RemoteWorkerHost scheduled task).

.EXAMPLE
.\scripts\deployment\Register-QCRemoteWorkerHostLogonConsole.ps1

.EXAMPLE
.\scripts\deployment\Register-QCRemoteWorkerHostLogonConsole.ps1 -Unregister
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$ShortcutName = 'QC Remote Worker Logs.lnk',

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = '',

    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = '',

    [Parameter(Mandatory = $false)]
    [switch]$AllWorkers,

    [Parameter(Mandatory = $false)]
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

$scriptsRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent $scriptsRoot
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$watchScript = Join-Path $scriptsRoot 'service\Watch-QCRemoteWorkerHostConsole.ps1'
if (-not (Test-Path -LiteralPath $watchScript)) {
    throw "Watch script not found: $watchScript"
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

$argList = New-Object System.Collections.Generic.List[string]
[void]$argList.Add('-NoExit')
[void]$argList.Add('-NoProfile')
[void]$argList.Add('-ExecutionPolicy')
[void]$argList.Add('Bypass')
[void]$argList.Add('-File')
[void]$argList.Add([string]$watchScript)
[void]$argList.Add('-AppSettingsPath')
[void]$argList.Add([string]$AppSettingsPath)
if ($AllWorkers.IsPresent) {
    [void]$argList.Add('-AllWorkers')
}

if ($PSCmdlet.ShouldProcess($shortcutPath, 'Create Startup shortcut')) {
    $wsh = New-Object -ComObject WScript.Shell
    $sc = $wsh.CreateShortcut($shortcutPath)
    $sc.TargetPath = 'powershell.exe'
    $sc.Arguments = ($argList -join ' ')
    $sc.WorkingDirectory = $RepoRoot
    $sc.WindowStyle = 1
    $sc.Description = 'Live tail of QC remote worker host job logs (read-only).'
    $sc.Save()

    Write-Host "Created Startup shortcut: $shortcutPath" -ForegroundColor Green
    Write-Host '  Opens at logon: Watch-QCRemoteWorkerHostConsole.ps1' -ForegroundColor Gray
    Write-Host '  Remove: .\scripts\deployment\Register-QCRemoteWorkerHostLogonConsole.ps1 -Unregister' -ForegroundColor Cyan
}
