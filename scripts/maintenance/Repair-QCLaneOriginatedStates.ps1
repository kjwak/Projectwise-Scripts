<#
.SYNOPSIS
Finds lane QC PDFs (*-prod/-chk/-rev) not at the expected post-intake state and optionally resets them.

.DESCRIPTION
Scans a ProjectWise Sheets folder for lane PDFs. Default target is qcWorkflow.states.readyForQc
(usually "Originated") — the state after successful initial QC prepend.

Default mode is preview only. Pass -ConfirmWrites to apply verified Set-PWDocumentState changes.
Also updates sheet_index / sheet_documents pw_state_name when database writes are enabled.

.EXAMPLE
.\scripts\maintenance\Repair-QCLaneOriginatedStates.ps1 `
  -FolderPath 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Estimating_Deliverables'

.EXAMPLE
.\scripts\maintenance\Repair-QCLaneOriginatedStates.ps1 `
  -FolderPath 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Estimating_Deliverables' `
  -LaneSuffix prod -ConfirmWrites
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$FolderPath,

    [string]$TargetState = '',
    [ValidateSet('prod', 'chk', 'rev', 'all')]
    [string]$LaneSuffix = 'prod',
    [string]$AppSettingsPath = '',
    [switch]$ConfirmWrites,
    [switch]$DryRun,
    [switch]$NoPrompt
)

$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
if (-not $scriptPath) {
    $scriptPath = Join-Path $PSScriptRoot 'Repair-QCLaneOriginatedStates.ps1'
}
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "MTA relaunch: could not resolve script path. Tried: $scriptPath"
    }
    $staMtaHelper = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'legacy\StaMtaRelaunch.ps1'
    if (-not (Test-Path -LiteralPath $staMtaHelper)) {
        throw "MTA relaunch helper not found: $staMtaHelper"
    }
    . $staMtaHelper
    $exeArgs = Build-PowerShellExeFileArgs -ScriptPath $scriptPath -BoundParameters $PSBoundParameters
    & powershell.exe @exeArgs
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Paths.psm1'
    'Core\Core.Telemetry.psm1'
    'Workflow\QC.Workflow.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.Discovery.psm1'
    'Database\Core.Database.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'ConvertTo-PWCmdletFolderPath'
    'Get-PWDocumentsInFolder'
    'Set-PWDocumentWorkflowStateVerified'
    'Invoke-PWAuthenticatedCommand'
) -Context 'Repair-QCLaneOriginatedStates bootstrap'

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

function _RQLO-NormalizeState {
    param([string]$StateName, [hashtable]$Config)
    $s = ([string]$StateName).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    if (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try { return [string](Format-QCWorkflowStateName -StateName $s -Config $Config) } catch { }
    }
    return $s
}

function _RQLO-StatesEqual {
    param([string]$A, [string]$B)
    return (([string]$A).Trim().ToLowerInvariant() -eq ([string]$B).Trim().ToLowerInvariant())
}

function _RQLO-TestLaneName {
    param([string]$DocumentName, [string]$LaneSuffix)
    if ([string]::IsNullOrWhiteSpace($DocumentName)) { return $false }
    if ($LaneSuffix -eq 'all') {
        return ($DocumentName -match '(?i)-(prod|chk|rev)\.pdf$')
    }
    return ($DocumentName -match ("(?i)-{0}\.pdf$" -f [regex]::Escape($LaneSuffix)))
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
if (-not $TargetState) {
    if ($config.ContainsKey('qcWorkflow') -and $config.qcWorkflow -and $config.qcWorkflow.states) {
        if ($config.qcWorkflow.states.readyForQc) {
            $TargetState = [string]$config.qcWorkflow.states.readyForQc
        } elseif ($config.qcWorkflow.states.qcReceived) {
            $TargetState = [string]$config.qcWorkflow.states.qcReceived
        }
    }
    if (-not $TargetState -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
        try {
            $wf = Get-QCWorkflowSettings -Config $config
            $TargetState = [string](Get-QCWorkflowStateName -Settings $wf -StateKey 'readyForQc')
        } catch { }
    }
    if (-not $TargetState) { $TargetState = 'Originated' }
}
$TargetState = _RQLO-NormalizeState -StateName $TargetState -Config $config

$doWrites = [bool]$ConfirmWrites -and -not [bool]$DryRun
if (-not $doWrites) {
    Write-Host 'PREVIEW MODE: no ProjectWise or database changes will be made. Pass -ConfirmWrites to apply.' -ForegroundColor Yellow
} else {
    Write-Host ("WRITE MODE: lane PDFs will be set to '{0}'." -f $TargetState) -ForegroundColor Green
}

$ds = [string]$config.projectWise.datasourceName
$credPath = [string]$config.projectWise.credentialPath
if ([string]::IsNullOrWhiteSpace($ds) -or [string]::IsNullOrWhiteSpace($credPath)) {
    throw 'projectWise.datasourceName or credentialPath missing in appsettings.'
}

$apiFolder = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
if ([string]::IsNullOrWhiteSpace($apiFolder)) { $apiFolder = $FolderPath }

$laneSuffixLocal = $LaneSuffix
$candidates = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -KeepSession -ScriptBlock {
    $docs = @(Get-PWDocumentsInFolder -FolderPath $using:apiFolder -JustThisFolder -ErrorAction Stop)
    $suffix = [string]$using:laneSuffixLocal
    $rows = @()
    foreach ($d in $docs) {
        $name = [string]$d.Name
        $isLane = $false
        if ($suffix -eq 'all') {
            $isLane = ($name -match '(?i)-(prod|chk|rev)\.pdf$')
        } else {
            $isLane = ($name -match ("(?i)-{0}\.pdf$" -f [regex]::Escape($suffix)))
        }
        if (-not $isLane) { continue }
        $guid = ''
        try { $guid = [string]$d.DocumentGUID } catch { }
        $state = ''
        foreach ($prop in @('WorkflowState', 'StateName', 'State', 'WorkflowStateName', 'CurrentState')) {
            try {
                $v = [string]$d.$prop
                if (-not [string]::IsNullOrWhiteSpace($v)) { $state = $v; break }
            } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($state) -and (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) {
            try {
                $state = [string](Get-PWDocumentWorkflowStateName -FolderPath $using:apiFolder -DocumentName $name -DocumentGuid $guid)
            } catch { }
        }
        $rows += [pscustomobject]@{
            documentName = $name
            documentGuid = $guid
            currentState = $state
        }
    }
    return ,$rows
}

$needsRepair = @($candidates | Where-Object {
    -not (_RQLO-StatesEqual -A $_.currentState -B $TargetState)
})

Write-Host ("Scanned {0} lane PDF(s); {1} need repair to '{2}'." -f $candidates.Count, $needsRepair.Count, $TargetState) -ForegroundColor Cyan
$needsRepair | Sort-Object documentName | Format-Table documentName, currentState, documentGuid -AutoSize

if ($needsRepair.Count -eq 0) {
    Write-Host 'Nothing to repair.' -ForegroundColor Green
    exit 0
}

if ($doWrites -and -not $NoPrompt -and $Host.Name -eq 'ConsoleHost') {
    $answer = Read-Host ("Apply {0} state write(s) to '{1}'? [y/N]" -f $needsRepair.Count, $TargetState)
    if ($answer -notmatch '^(?i)y(es)?$') {
        Write-Host 'Aborted.' -ForegroundColor Yellow
        exit 2
    }
}

if (-not $doWrites) {
    Write-Host 'Preview complete.' -ForegroundColor Green
    exit 0
}

$ok = 0
$fail = 0
foreach ($row in $needsRepair) {
    if (-not $PSCmdlet.ShouldProcess($row.documentName, "Set workflow state to $TargetState")) { continue }
    try {
        $write = Set-PWDocumentWorkflowStateVerified -Config $config `
            -FolderPath $apiFolder -DocumentName $row.documentName -DocumentGuid $row.documentGuid `
            -TargetState $TargetState -QcProcessType 'production' -IsLaneAuthority:$true -WriteScope 'lane'
        if (-not $write.verified) {
            throw ("unverified write (read-back='{0}')" -f $write.readBackState)
        }
        if ($row.documentGuid -and (Get-Command -Name 'Update-QCSheetIndexPwStateName' -ErrorAction SilentlyContinue)) {
            try {
                [void](Update-QCSheetIndexPwStateName -Config $config -DocumentGuid $row.documentGuid -PwStateName $TargetState)
            } catch { }
        }
        Write-Host ("OK  {0}: {1} -> {2}" -f $row.documentName, $row.currentState, $TargetState) -ForegroundColor Green
        $ok++
    } catch {
        Write-Host ("FAIL {0}: {1}" -f $row.documentName, $_.Exception.Message) -ForegroundColor Red
        $fail++
    }
}

Write-Host ("Done. succeeded={0} failed={1}" -f $ok, $fail) -ForegroundColor $(if ($fail -gt 0) { 'Yellow' } else { 'Green' })
if ($fail -gt 0) { exit 1 }
exit 0
