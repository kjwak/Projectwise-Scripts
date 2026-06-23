# Unit tests for Get-QCEffectiveDryRunPolicy (layered dry-run side effects).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core/Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/Core/Core.Runtime.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

$prod = @{
    dryRun = $false
    processors = @{ dryRun = @{ invokeHandler = $true; allowStateChange = $false } }
    qcWorkflow = @{ dryRunWriteback = $false }
    notifications = @{ enabled = $true; dryRun = $false }
    database = @{ enabled = $true; allowWritesInDryRun = $false }
    statusSet = @{ writeBackToPW = $true }
}
$p = Get-QCEffectiveDryRunPolicy -Config $prod -Role 'worker'
Assert-Eq $p.effectivePolicy.enqueueJobs $null 'worker policy has no enqueueJobs'
Assert-Eq $p.effectivePolicy.lockAndMoveQueueJobs $true 'production worker locks queue'
Assert-Eq $p.effectivePolicy.writePwFilesViaProcessors $true 'production worker may write PW'
Assert-Eq $p.effectivePolicy.sendNotificationEmail $true 'production sends email'

$w = Get-QCEffectiveDryRunPolicy -Config $prod -Role 'watcher'
Assert-Eq $w.effectivePolicy.enqueueJobs $true 'production watcher enqueues'
Assert-Eq $w.effectivePolicy.writePwFilesViaProcessors $false 'watcher never writes PW via processors'

$test = @{
    dryRun = $true
    processors = @{ dryRun = @{ invokeHandler = $true; allowStateChange = $true } }
    qcWorkflow = @{ dryRunWriteback = $true }
    notifications = @{ enabled = $false; dryRun = $true }
    database = @{ enabled = $false; allowWritesInDryRun = $false }
}
$t = Get-QCEffectiveDryRunPolicy -Config $test -Role 'worker'
Assert-Eq $t.effectivePolicy.enqueueJobs $null 'worker: no enqueue key'
Assert-Eq $t.effectivePolicy.invokeProcessorHandlers $true 'test profile invokes handlers under dry run'
Assert-Eq $t.effectivePolicy.writePwFilesViaProcessors $false 'global dry run blocks PW file writes'
Assert-Eq $t.effectivePolicy.writePwWorkflowAttributes $false 'dryRunWriteback blocks workflow writes'
Assert-Eq $t.effectivePolicy.writeSqlTelemetry $false 'database disabled'

$tw = Get-QCEffectiveDryRunPolicy -Config $test -Role 'watcher'
Assert-Eq $tw.effectivePolicy.enqueueJobs $false 'global dry run suppresses enqueue'

Write-Host 'OK test_effective_dry_run_policy.ps1'
