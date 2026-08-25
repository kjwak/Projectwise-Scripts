<#
.SYNOPSIS
Registers a Windows Scheduled Task to run Start-QCPipelineDashboard.ps1 at machine startup.

.DESCRIPTION
Runs the QC server dashboard (watcher + local workers) as the specified user even when
that user is not logged on (LogonType Password). Adds a startup delay so SQL, ProjectWise,
and local disks are usually ready after reboot.

The terminal UI is not visible in Session 0. Use Register-QCPipelineDashboardLogonConsole.ps1
for the logon ops GUI. Do not register this on the modelling PC.

Requires elevation to register. You will be prompted for the run-as account password
unless -Credential is supplied. After register, the run-as account is granted Enable/Disable
on the task so the unelevated ops GUI can toggle Pipeline on/off. For a task that already
exists, re-run elevated with -GrantOperatorAccess (no password / no recreate).

.EXAMPLE
.\scripts\deployment\Register-QCPipelineDashboardTask.ps1

.EXAMPLE
.\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -GrantOperatorAccess

.EXAMPLE
.\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -RepoRoot 'C:\Users\jflint\Documents\github\Prepend PDF QC'
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TaskName = 'QC-PipelineDashboard',

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
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$Unregister,

    [Parameter(Mandatory = $false)]
    [switch]$GrantOperatorAccess
)

$ErrorActionPreference = 'Stop'

$expectedHost = 'PXBENTLEY01'
$modellingHost = 'AZTEC002799'
$prodCloneRoot = 'C:\Users\jflint\Documents\github\Prepend PDF QC'

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

function _Quote-TaskArg {
    param([Parameter(Mandatory)][string]$Value)
    if ($Value -notmatch '[ \t"]') { return $Value }
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

if (-not (_Test-IsElevated)) {
    throw 'Register-QCPipelineDashboardTask must run elevated (Run as administrator).'
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

$dashScript = Join-Path $RepoRoot 'scripts\service\Start-QCPipelineDashboard.ps1'
if (-not (Test-Path -LiteralPath $dashScript)) {
    throw "Dashboard script not found: $dashScript"
}

if ($Unregister.IsPresent) {
    if ($PSCmdlet.ShouldProcess($TaskName, 'Unregister scheduled task')) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Unregistered scheduled task: $TaskName" -ForegroundColor Green
    }
    return
}

$hostName = [string]$env:COMPUTERNAME
if (-not $Force.IsPresent) {
    if ($hostName -ieq $modellingHost) {
        throw "Refusing to register the full pipeline dashboard on modelling host $hostName. Use Register-QCRemoteWorkerHostTask.ps1 there, or pass -Force to override."
    }
    if ($hostName -ine $expectedHost) {
        throw "This task is for QC server $expectedHost (this host is $hostName). RDP to $expectedHost and re-run, or pass -Force to override."
    }
}

if ($GrantOperatorAccess.IsPresent) {
    if ([string]::IsNullOrWhiteSpace($UserName)) {
        $UserName = [string]$env:USERDOMAIN + '\' + [string]$env:USERNAME
    }
    $aclHelper = Join-Path $PSScriptRoot 'Set-QCScheduledTaskOperatorAcl.ps1'
    & $aclHelper -TaskName $TaskName -UserName $UserName
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
if (-not (Test-Path -LiteralPath $AppSettingsPath)) {
    throw "AppSettings not found: $AppSettingsPath"
}
$AppSettingsPath = (Resolve-Path -LiteralPath $AppSettingsPath).Path

$argList = New-Object System.Collections.Generic.List[string]
[void]$argList.Add('-NoProfile')
[void]$argList.Add('-ExecutionPolicy')
[void]$argList.Add('Bypass')
[void]$argList.Add('-MTA')
[void]$argList.Add('-File')
[void]$argList.Add((_Quote-TaskArg -Value $dashScript))
[void]$argList.Add('-AppSettingsPath')
[void]$argList.Add((_Quote-TaskArg -Value $AppSettingsPath))

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument ($argList -join ' ') `
    -WorkingDirectory $RepoRoot

$trigger = New-ScheduledTaskTrigger -AtStartup
if ($StartupDelaySeconds -gt 0) {
    $trigger.Delay = _ConvertTo-Iso8601DurationSeconds -Seconds $StartupDelaySeconds
}

# ExecutionTimeLimit 0 disables the default 72-hour stop so the dashboard can run continuously.
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -MultipleInstances IgnoreNew

$desc = 'QC pipeline dashboard (watcher + workers). Starts at boot whether the user is logged on or not. Terminal UI is not visible in Session 0; use the logon log console.'

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

    $aclHelper = Join-Path $PSScriptRoot 'Set-QCScheduledTaskOperatorAcl.ps1'
    try {
        & $aclHelper -TaskName $TaskName -UserName $UserName
    } catch {
        Write-Host ("Warning: could not grant unelevated Pipeline on/off to {0}: {1}" -f $UserName, $_.Exception.Message) -ForegroundColor Yellow
        Write-Host 'Re-run elevated: .\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -GrantOperatorAccess' -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host "Registered: $TaskName" -ForegroundColor Green
    Write-Host "  User:     $UserName (runs when logged off)" -ForegroundColor Gray
    Write-Host "  Trigger:  At startup (+${StartupDelaySeconds}s delay)" -ForegroundColor Gray
    Write-Host "  Script:   $dashScript" -ForegroundColor Gray
    Write-Host "  Working:  $RepoRoot" -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Session 0 has no desktop. Open a read-only log window at logon with:' -ForegroundColor Yellow
    Write-Host '  .\scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1' -ForegroundColor Cyan
    Write-Host 'Do not start a second interactive dashboard on this host (singleton lock / duplicate watchers).' -ForegroundColor Yellow
    Write-Host ''
    Write-Host ("Test now:  Start-ScheduledTask -TaskName '{0}'" -f $TaskName) -ForegroundColor Cyan
    Write-Host 'Remove:    .\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -Unregister' -ForegroundColor Cyan
}
