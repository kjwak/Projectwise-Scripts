# Exercise the service entrypoints up to their configuration read in fresh 5.1 sessions.
[CmdletBinding()]
param(
    [ValidateSet('Watcher', 'Worker')]
    [string]$Service,
    [switch]$PublicationFirst
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$bootstrapRepo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

if (-not $Service) {
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    foreach ($entry in @('Watcher', 'Worker')) {
        foreach ($preload in @($false, $true)) {
            $childArguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                '-File', $PSCommandPath, '-Service', $entry)
            if ($preload) { $childArguments += '-PublicationFirst' }
            & $windowsPowerShell @childArguments
            if ($LASTEXITCODE -ne 0) {
                throw "$entry bootstrap failed (publication-first=$preload, exit $LASTEXITCODE)"
            }
        }
    }
    return
}

$env:PSModulePath = Join-Path $PSHOME 'Modules'
$script:BootstrapExternalCalls = [System.Collections.Generic.List[string]]::new()
foreach ($externalCommand in @('Open-PWConnection', 'New-PWLogin', 'Get-PWDocumentsBySearch', 'Start-ScheduledTask')) {
    Set-Item -LiteralPath "Function:global:$externalCommand" -Value {
        $script:BootstrapExternalCalls.Add($MyInvocation.MyCommand.Name)
        throw "Unexpected external operation: $($MyInvocation.MyCommand.Name)"
    }
}

# This standalone path must be absent so the real config reader exits before service work.
$missingSettings = Join-Path ([System.IO.Path]::GetTempPath()) ("pwqc-bootstrap-" + [guid]::NewGuid() + ".json")
$expectedStop = "appsettings.json not found: $missingSettings"
$bootstrapEntry = @{
    Watcher = 'scripts\service\Watch-QCTrigger.ps1'
    Worker  = 'scripts\service\Run-QCProcessor.ps1'
}
$bootstrapPath = Join-Path $bootstrapRepo $bootstrapEntry[$Service]
$bootstrapModules = (Join-Path $bootstrapRepo 'modules') + '\'
$requiredExports = @('Read-QCAppSettings', 'Write-QCJsonLog', 'Get-Sha256TextHex', 'Test-QCDatabaseEnabled')
if ($Service -eq 'Watcher') {
    $requiredExports += @(
        'Connect-PW', 'Get-PWCredentialFromFile', 'Ensure-PWDiscoveryModuleLoaded',
        'Get-PWDocumentsInFolder', 'Get-PWImmediateChildFolders', 'Invoke-AuditTrailScan',
        'Invoke-StatusSetScheduledPublication'
    )
} else {
    $requiredExports += @('Get-NextQCJob', 'Invoke-QCPrependProcessor', 'Invoke-QCProcessorByType')
}

function Assert-BootstrapExports([string]$Stage) {
    foreach ($exportName in $requiredExports) {
        $command = Get-Command -Name $exportName -CommandType Function -ErrorAction SilentlyContinue
        if (-not $command.Module.Path -or -not $command.Module.Path.StartsWith($bootstrapModules, [StringComparison]::OrdinalIgnoreCase)) {
            throw "$Stage - missing repository module export $exportName"
        }
    }
    if ($script:BootstrapExternalCalls.Count -ne 0) {
        throw "$Stage - unexpected external operations: $($script:BootstrapExternalCalls -join ', ')"
    }
    Write-Host "PASS $Stage ($($requiredExports.Count) exports)"
}

if ($PublicationFirst) {
    Import-Module (Join-Path $bootstrapRepo 'modules\Processing\QC.StatusSetPublication.psm1') -Force
    if ($Service -eq 'Worker') { $requiredExports += 'Invoke-StatusSetScheduledPublication' }
}

$bootstrapRuns = if ($Service -eq 'Watcher') { 1 } else { 1..2 }
foreach ($bootstrapRun in $bootstrapRuns) {
    $stage = "$Service startup $bootstrapRun (publication-first=$($PublicationFirst.IsPresent))"
    try {
        # Keep script-scoped watcher exports and restore functions visible to the assertions.
        . $bootstrapPath -AppSettingsPath $missingSettings
        throw "$stage did not stop at the configuration read"
    } catch {
        if ($_.Exception.Message -ne $expectedStop) { throw }
    }
    Assert-BootstrapExports -Stage $stage
}

if ($Service -eq 'Watcher') {
    _Watch-ReloadWatchModules
    Assert-BootstrapExports -Stage "$stage reload and restore"
}
