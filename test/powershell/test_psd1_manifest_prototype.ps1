# Phase 4G prototype: QC.Core.psd1 and QC.Queue.psd1 manifest smoke tests (test-only; not production).
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulesRoot = Join-Path $repoRoot 'modules'
$coreManifest = Join-Path $modulesRoot 'Core\QC.Core.psd1'
$queueManifest = Join-Path $modulesRoot 'Queue\QC.Queue.psd1'
$fail = 0

function Assert-True([bool]$Cond, [string]$Msg) {
    if (-not $Cond) {
        Write-Host "FAIL: $Msg" -ForegroundColor Red
        $script:fail++
    } else {
        Write-Host "OK:   $Msg" -ForegroundColor Green
    }
}

function Assert-Command([string]$Name, [string]$Msg) {
    Assert-True ($null -ne (Get-Command $Name -ErrorAction SilentlyContinue)) $Msg
}

Write-Host '=== Manifest files ===' -ForegroundColor Cyan
Assert-True (Test-Path -LiteralPath $coreManifest) 'QC.Core.psd1 exists'
Assert-True (Test-Path -LiteralPath $queueManifest) 'QC.Queue.psd1 exists'

Write-Host '=== Test-ModuleManifest ===' -ForegroundColor Cyan
$coreInfo = Test-ModuleManifest -Path $coreManifest -ErrorAction Stop
$queueInfo = Test-ModuleManifest -Path $queueManifest -ErrorAction Stop
Assert-True ($coreInfo.NestedModules.Count -eq 8) 'QC.Core lists 8 nested modules'
Assert-True ($queueInfo.NestedModules.Count -eq 5) 'QC.Queue lists 5 nested modules'

Write-Host '=== Import QC.Core.psd1 (PowerShell 5.1 session) ===' -ForegroundColor Cyan
Import-Module $coreManifest -Force -WarningAction SilentlyContinue
$coreProbes = @(
    'New-QCResult'
    'Get-QCTimestamp'
    'Normalize-QCPath'
    'Read-QCAppSettings'
    'Write-QCJsonLog'
    'Get-Sha256TextHex'
    'Write-QCAutomationEvent'
    'Get-QCWatcherMode'
)
foreach ($cmd in $coreProbes) {
    Assert-Command $cmd "QC.Core exposes $cmd"
}

Write-Host '=== Import QC.Queue.psd1 after QC.Core ===' -ForegroundColor Cyan
Import-Module $queueManifest -Force -WarningAction SilentlyContinue
$queueProbes = @(
    'Get-NextQCJob'
    'New-QCJobObject'
    'Move-QCJobWithLockRetries'
    'Test-QCPathAllowed'
    'Resolve-QCTriggerMatch'
)
foreach ($cmd in $queueProbes) {
    Assert-Command $cmd "QC.Queue exposes $cmd"
}

Write-Host '=== Folder implementation after manifest import (Phase 4H) ===' -ForegroundColor Cyan
$flatResults = Join-Path $modulesRoot 'Core.Results.psm1'
$folderResults = Join-Path $modulesRoot 'Core\Core.Results.psm1'
Assert-True (-not (Test-Path -LiteralPath $flatResults)) 'flat Core.Results shim removed'
Assert-True (Test-Path -LiteralPath $folderResults) 'folder Core.Results implementation exists'
Import-Module $folderResults -Force -WarningAction SilentlyContinue
Assert-Command 'New-QCResult' 'folder path import resolves New-QCResult'
Assert-Command 'New-QCSuccessResult' 'folder path import resolves New-QCSuccessResult'

Write-Host '=== Manifest does not require live PW/SQL/Graph (import-only) ===' -ForegroundColor Cyan
Assert-True (-not (Get-Command Connect-PW -ErrorAction SilentlyContinue)) 'Connect-PW not loaded by core+queue manifests alone'
Assert-True (-not (Get-Command Invoke-QCNotificationForStateChange -ErrorAction SilentlyContinue)) 'notification cmd not loaded by manifests alone'

if ($fail -gt 0) {
    Write-Host "test_psd1_manifest_prototype.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_psd1_manifest_prototype.ps1 passed' -ForegroundColor Green
