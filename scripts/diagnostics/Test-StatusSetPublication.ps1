<#
.SYNOPSIS
Manually preview or publish configured ProjectWise Status Set archive copies.

.DESCRIPTION
Calls the same publication functions the Monday 06:00 reconciliation uses
(Invoke-StatusSetPublication). Preview is the default and does not create
ProjectWise documents.

-RunReconciliation runs Invoke-StatusSetNativeJob for the configured source
folders only. A full Watch-QCTrigger reconciliation also advances the audit
watermark and marks full_scan_schedule slots complete for every watched
project, so this script does not start that pass. The native status-set
processor is the component reconciliation enqueues. This script does not
write full-scan schedule state or the audit watermark.

-UseLatestExisting publishes the current source PDFs even when they were not
produced by the Monday cycle. The report shows each source date and does not
mark those files as newly generated. Scheduled runs never set this switch.

.EXAMPLE
.\scripts\diagnostics\Test-StatusSetPublication.ps1 -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate '2026-11-02' -UseLatestExisting -Mode Preview

.EXAMPLE
.\scripts\diagnostics\Test-StatusSetPublication.ps1 -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate '2026-11-02' -UseLatestExisting -Mode Publish

.EXAMPLE
.\scripts\diagnostics\Test-StatusSetPublication.ps1 -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate '2026-11-02' -RunReconciliation -Mode Publish
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ProjectKey,

    [Parameter(Mandatory)]
    [datetime]$AsOfDate,

    [ValidateSet('Preview', 'Publish')]
    [string]$Mode = 'Preview',

    [switch]$UseLatestExisting,

    [switch]$RunReconciliation,

    [string]$AppSettingsPath = ''
)

$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'Test-StatusSetPublication.ps1' }
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
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
    'Processing\QC.StatusSet.psm1'
    'Processing\QC.StatusSetPublication.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Write-QCJsonLog'
    'Invoke-StatusSetPublication'
    'Invoke-StatusSetPublicationSourceReconcile'
) -Context 'Test-StatusSetPublication bootstrap'
# Foundation restore drops PW.Connection session exports. Re-import it last.
Import-QCModuleGlobal -RelativePath 'ProjectWise\PW.Connection.psm1'
if (-not (Get-Command -Name 'Connect-PW' -ErrorAction SilentlyContinue)) {
    throw 'Test-StatusSetPublication bootstrap failed: missing command(s): Connect-PW'
}

function Write-StatusSetPublicationReport {
    param($Result, [string]$ProjectKey, [datetime]$AsOfDate, [string]$Mode, [bool]$UseLatestExisting, [bool]$ScheduledEnabled, $Reconcile)
    Write-Host ''
    Write-Host 'PROJECTWISE STATUS SET PUBLICATION TEST'
    Write-Host ''
    Write-Host ("Project: {0}" -f $ProjectKey)
    Write-Host ("Mode: {0}" -f $Mode)
    Write-Host ("Date: {0}" -f $AsOfDate.ToString('yyyy-MM-dd'))
    Write-Host ("Run: {0}" -f 'Manual')
    if ($UseLatestExisting) {
        Write-Host 'Provenance bypass: UseLatestExisting (source dates below are the documents that were found; they are not marked newly generated)'
    }
    if ($ScheduledEnabled) {
        Write-Host 'Scheduled publication: ENABLED'
    } else {
        Write-Host 'Scheduled publication: DISABLED (statusSetPublication.enabled is false; this manual run does not change that)'
    }
    if ($Reconcile) {
        Write-Host 'Reconciliation: existing Invoke-StatusSetNativeJob for configured folders only; schedule state was not modified'
        foreach ($row in @($Reconcile.Data.sets)) {
            $gen = if ([bool]$row.newlyGenerated) { 'newly generated' } else { 'not newly generated' }
            Write-Host ("  {0} reconcile: {1} ({2})" -f [string]$row.id, [string]$row.reconcileCode, $gen)
        }
    }
    Write-Host ''
    if (-not $Result -or -not $Result.Data -or -not $Result.Data.segments) {
        Write-Host ("Overall Result: FAILURE")
        if ($Result -and $Result.Message) { Write-Host $Result.Message }
        return
    }
    foreach ($seg in @($Result.Data.segments)) {
        Write-Host ([string]$seg.id)
        Write-Host ("  Source:       {0}" -f [string]$seg.sourceStatus)
        $srcDate = 'unknown'
        if ($seg.sourceLastWriteUtc) {
            try { $srcDate = ([datetime]$seg.sourceLastWriteUtc).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } catch { $srcDate = [string]$seg.sourceLastWriteUtc }
        }
        Write-Host ("  SourceFile:   {0}" -f [string]$seg.sourceName)
        Write-Host ("  SourceDate:   {0}" -f $srcDate)
        Write-Host ("  SourceSize:   {0}" -f [int64]$seg.sourceSize)
        Write-Host ("  Provenance:   {0}" -f $(if ($seg.provenance) { [string]$seg.provenance } else { 'n/a' }))
        Write-Host ("  NewlyGenerated: {0}" -f $(if ([bool]$seg.newlyGenerated) { 'YES' } else { 'NO' }))
        Write-Host ("  Validation:   {0}" -f [string]$seg.validation)
        if ($seg.validationDetail) { Write-Host ("  Detail:       {0}" -f [string]$seg.validationDetail) }
        $dest = if ($seg.destinationName) { [string]$seg.destinationName } else { '(not resolved)' }
        Write-Host ("  Destination:  {0}" -f $dest)
        if ($seg.destinationFolder) { Write-Host ("  DestFolder:   {0}" -f [string]$seg.destinationFolder) }
        if ([bool]$seg.conflict) { Write-Host '  Conflict:     YES' }
        Write-Host ("  Copy:         {0}" -f [string]$seg.copy)
        Write-Host ("  Verification: {0}" -f [string]$seg.verification)
        if ($Mode -eq 'Preview') {
            Write-Host ("  Publishable:  {0}" -f $(if ([bool]$seg.canPublish) { 'YES' } else { 'NO' }))
        }
        if ($seg.message) { Write-Host ("  Message:      {0}" -f [string]$seg.message) }
        Write-Host ''
    }
    $overall = if ($Result.Data.overall) { [string]$Result.Data.overall } else { 'FAILURE' }
    Write-Host ("Overall Result: {0}" -f $overall)
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) {
    Write-Host $cfgRes.Message -ForegroundColor Red
    exit 2
}
$config = [hashtable]$cfgRes.Data.config

$settingsRes = Get-StatusSetPublicationSettings -Config $config
if (-not $settingsRes.IsSuccess) {
    Write-Host $settingsRes.Message -ForegroundColor Red
    exit 2
}
$projectMatch = @($settingsRes.Data.projects | Where-Object { [string]$_.projectKey -ieq $ProjectKey })
if ($projectMatch.Count -eq 0) {
    Write-Host "Project '$ProjectKey' has no status set publication configuration." -ForegroundColor Red
    exit 2
}

$reconcileResult = $null
$hints = $null
if ($RunReconciliation.IsPresent) {
    $reconcileResult = Invoke-StatusSetPublicationSourceReconcile -Config $config -ProjectKey $ProjectKey
    if (-not $reconcileResult.IsSuccess -and $reconcileResult.Code -eq 'STATUS_SET_PUBLISH_CONFIG') {
        Write-StatusSetPublicationReport -Result $reconcileResult -ProjectKey $ProjectKey -AsOfDate $AsOfDate -Mode $Mode -UseLatestExisting:([bool]$UseLatestExisting.IsPresent) -ScheduledEnabled:([bool]$settingsRes.Data.enabled) -Reconcile $null
        exit 2
    }
    if ($reconcileResult.Data -and $reconcileResult.Data.hints) {
        $hints = $reconcileResult.Data.hints
    }
}

$pw = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $pwNorm = ConvertTo-HashtableDeep -Value $config.projectWise
    if ($pwNorm) { $pw = $pwNorm }
}
$ds = if ($pw.ContainsKey('datasourceName') -and $pw.datasourceName) { [string]$pw.datasourceName } else { '' }
$credPath = if ($pw.ContainsKey('credentialPath') -and $pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
if ([string]::IsNullOrWhiteSpace($ds)) {
    Write-Host 'projectWise.datasourceName is not configured.' -ForegroundColor Red
    exit 2
}

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) {
    Write-Host ($credRes.Code + ': ' + $credRes.Message) -ForegroundColor Red
    exit 1
}
$conn = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $conn.IsSuccess) {
    Write-Host ($conn.Code + ': ' + $conn.Message) -ForegroundColor Red
    exit 1
}

$result = $null
try {
    $pubParams = @{
        Config = $config
        ProjectKey = $ProjectKey
        AsOfDate = $AsOfDate.Date
        Mode = $Mode
        RunKind = 'Manual'
        CheckInFlight = $true
    }
    if ($UseLatestExisting.IsPresent) { $pubParams.UseLatestExisting = $true }
    if ($hints) { $pubParams.ProvenanceHints = $hints }
    if ($RunReconciliation.IsPresent) { $pubParams.CycleStartUtc = [datetime]::UtcNow }
    $result = Invoke-StatusSetPublication @pubParams
} finally {
    if (Get-Command -Name 'Disconnect-PW' -ErrorAction SilentlyContinue) {
        try { Disconnect-PW | Out-Null } catch { }
    }
}

Write-StatusSetPublicationReport -Result $result -ProjectKey $ProjectKey -AsOfDate $AsOfDate -Mode $Mode -UseLatestExisting:([bool]$UseLatestExisting.IsPresent) -ScheduledEnabled:([bool]$settingsRes.Data.enabled) -Reconcile $reconcileResult

if (-not $result -or -not $result.IsSuccess -or [string]$result.Data.overall -ne 'SUCCESS') {
    exit 1
}
exit 0
