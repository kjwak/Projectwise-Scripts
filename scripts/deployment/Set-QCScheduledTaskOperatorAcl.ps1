<#
.SYNOPSIS
Grants a Windows account Enable/Disable/Start/Stop on one scheduled task.

.DESCRIPTION
Task Scheduler default DACL on tasks created from an elevated registrar is
Administrators-only. The logon ops GUI runs unelevated, so Pipeline on/off
fails with Access is denied until this ACE exists.

Requires elevation. Safe to re-run; no-ops if the SID is already on the DACL.

.EXAMPLE
.\scripts\deployment\Set-QCScheduledTaskOperatorAcl.ps1
.\scripts\deployment\Set-QCScheduledTaskOperatorAcl.ps1 -TaskName QC-PipelineDashboard -UserName 'TYPSA\jflint'
#>
[CmdletBinding()]
param(
    [string]$TaskName = 'QC-PipelineDashboard',
    [string]$UserName = ''
)

$ErrorActionPreference = 'Stop'

function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsElevated)) {
    throw 'Set-QCScheduledTaskOperatorAcl must run elevated (Run as administrator).'
}

if ([string]::IsNullOrWhiteSpace($UserName)) {
    $UserName = [string]$env:USERDOMAIN + '\' + [string]$env:USERNAME
}

$sid = $null
try {
    $sid = (New-Object System.Security.Principal.NTAccount($UserName)).Translate(
        [System.Security.Principal.SecurityIdentifier]
    ).Value
} catch {
    throw ("Could not resolve account '{0}' to a SID: {1}" -f $UserName, $_.Exception.Message)
}

$svc = New-Object -ComObject 'Schedule.Service'
$svc.Connect()
$folder = $svc.GetFolder('\')
$task = $null
try { $task = $folder.GetTask($TaskName) } catch { $task = $null }
if (-not $task) {
    throw ("Scheduled task '{0}' was not found." -f $TaskName)
}

$sddl = ''
try { $sddl = [string]$task.GetSecurityDescriptor(4) } catch { $sddl = '' }
if ([string]::IsNullOrWhiteSpace($sddl)) {
    try { $sddl = [string]$task.GetSecurityDescriptor(7) } catch { $sddl = '' }
}
if ([string]::IsNullOrWhiteSpace($sddl)) {
    throw ("Could not read security descriptor for task '{0}'." -f $TaskName)
}

if ($sddl -match [regex]::Escape($sid)) {
    Write-Host ("Operator access already present on {0} for {1}" -f $TaskName, $UserName) -ForegroundColor Green
    return
}

$ace = "(A;;FA;;;$sid)"
$daclAt = $sddl.IndexOf('D:')
if ($daclAt -ge 0) {
    $after = $sddl.Substring($daclAt + 2)
    $flags = ''
    if ($after -match '^([A-Z]+)') { $flags = $Matches[1] }
    $sddl = $sddl.Insert($daclAt + 2 + $flags.Length, $ace)
} else {
    $sddl = 'D:' + $ace + $sddl
}
$task.SetSecurityDescriptor($sddl, 0)
Write-Host ("Granted Enable/Disable/Start/Stop on {0} to {1}" -f $TaskName, $UserName) -ForegroundColor Green
