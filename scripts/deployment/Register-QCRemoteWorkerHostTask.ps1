<#
.SYNOPSIS
Registers a Windows Scheduled Task to run Start-QCRemoteWorkerHost.ps1 at machine startup.

.DESCRIPTION
Runs the remote worker supervisor as the specified user even when that user is not
logged on (LogonType Password). Adds a startup delay so UNC queue and network are
usually ready after reboot.

Requires elevation to register. You will be prompted for the run-as account password
unless -Credential is supplied.

.EXAMPLE
.\scripts\deployment\Register-QCRemoteWorkerHostTask.ps1

.EXAMPLE
.\scripts\deployment\Register-QCRemoteWorkerHostTask.ps1 -StartupDelaySeconds 120 -Credential (Get-Credential)
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TaskName = 'QC-RemoteWorkerHost',

    [Parameter(Mandatory = $false)]
    [string]$UserName = '',

    [Parameter(Mandatory = $false)]
    [pscredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = '',

    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = '',

    [Parameter(Mandatory = $false)]
    [int]$StartupDelaySeconds = 90,

    [Parameter(Mandatory = $false)]
    [switch]$AllowUncQueue,

    [Parameter(Mandatory = $false)]
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

function _ConvertTo-Iso8601DurationSeconds {
    param([Parameter(Mandatory)][int]$Seconds)
    if ($Seconds -le 0) { return $null }
    return ('PT{0}S' -f $Seconds)
}

function _Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (_Test-IsElevated)) {
    throw 'Register-QCRemoteWorkerHostTask must run elevated (Run as administrator).'
}

$scriptsRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent $scriptsRoot
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$hostScript = Join-Path $scriptsRoot 'service\Start-QCRemoteWorkerHost.ps1'
if (-not (Test-Path -LiteralPath $hostScript)) {
    throw "Remote worker host script not found: $hostScript"
}

if ($Unregister.IsPresent) {
    if ($PSCmdlet.ShouldProcess($TaskName, 'Unregister scheduled task')) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Unregistered scheduled task: $TaskName" -ForegroundColor Green
    }
    return
}

if ($Credential) {
    $UserName = [string]$Credential.UserName
    $passwordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
} else {
    if ([string]::IsNullOrWhiteSpace($UserName)) {
        $UserName = [string]$env:USERDOMAIN + '\' + [string]$env:USERNAME
    }
    Write-Host "Enter the password for $UserName (stored encrypted with the task)." -ForegroundColor Cyan
    $Credential = Get-Credential -UserName $UserName -Message 'Password for run-whether-logged-on task'
    $UserName = [string]$Credential.UserName
    $passwordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
}

if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $RepoRoot 'appsettings.json'
}

$argList = New-Object System.Collections.Generic.List[string]
[void]$argList.Add('-NoProfile')
[void]$argList.Add('-ExecutionPolicy')
[void]$argList.Add('Bypass')
[void]$argList.Add('-File')
[void]$argList.Add([string]$hostScript)
[void]$argList.Add('-AppSettingsPath')
[void]$argList.Add([string]$AppSettingsPath)
if ($AllowUncQueue.IsPresent) {
    [void]$argList.Add('-AllowUncQueue')
}

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument ($argList -join ' ') `
    -WorkingDirectory $RepoRoot

$trigger = New-ScheduledTaskTrigger -AtStartup
if ($StartupDelaySeconds -gt 0) {
    $trigger.Delay = _ConvertTo-Iso8601DurationSeconds -Seconds $StartupDelaySeconds
}

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -MultipleInstances IgnoreNew

$desc = 'QC remote worker host (processor-only). Starts at boot; claims QC_PREPEND from shared JSON queue.'

if ($PSCmdlet.ShouldProcess($TaskName, 'Register scheduled task')) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    try {
        # -User/-Password and -Principal are different Register-ScheduledTask parameter sets; use User+Password for logon-when-logged-off.
        Register-ScheduledTask `
            -TaskName $TaskName `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Description $desc `
            -User $UserName `
            -Password $passwordPlain -ErrorAction Stop | Out-Null
    } catch {
        throw "Failed to register scheduled task '$TaskName': $($_.Exception.Message)"
    }

    $verify = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $verify) {
        throw "Scheduled task '$TaskName' was not created. Re-run from an elevated PowerShell window."
    }

    Write-Host ''
    Write-Host "Registered: $TaskName" -ForegroundColor Green
    Write-Host "  User:     $UserName (runs when logged off)" -ForegroundColor Gray
    Write-Host "  Trigger:  At startup (+${StartupDelaySeconds}s delay)" -ForegroundColor Gray
    Write-Host "  Script:   $hostScript" -ForegroundColor Gray
    Write-Host "  Working:  $RepoRoot" -ForegroundColor Gray
    if ($AllowUncQueue.IsPresent) { Write-Host '  Flags:    -AllowUncQueue' -ForegroundColor Gray }
    Write-Host ''
    Write-Host ("Test now:  Start-ScheduledTask -TaskName '{0}'" -f $TaskName) -ForegroundColor Cyan
    Write-Host 'Remove:    .\scripts\deployment\Register-QCRemoteWorkerHostTask.ps1 -Unregister' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'If prepend fails after reboot (PW/UNC), confirm jflint can reach the queue share and pw_cred.txt from a non-interactive session.' -ForegroundColor Yellow
}
