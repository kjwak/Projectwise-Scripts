<#
.SYNOPSIS
Resolve read-only diagnostics locations for the live QC queue/log share.

.DESCRIPTION
Diagnostics-only. Does not change production queue.rootDir, watcher/worker log
paths, or MCP behavior. Never writes to the queue or logs.

Precedence for QueueRoot:
  1. -QueueRoot
  2. $env:QC_DEBUG_QUEUE_ROOT
  3. config/qc-debug-locations.local.json -> queueRoot
  4. Default UNC: \\192.168.22.90\QC_Queue

Precedence for LogDir:
  1. -LogDir
  2. $env:QC_DEBUG_LOG_DIR
  3. local config logDir
  4. Derived: Join-Path QueueRoot '_logs'

Optional -UseAppSettingsQueueRoot reads production appsettings queue.rootDir
(escape hatch for old local-queue debugging only).
#>

$script:QCDebugDefaultQueueRoot = '\\192.168.22.90\QC_Queue'

function Resolve-QCDebugLocations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$QueueRoot,

        [Parameter(Mandatory = $false)]
        [string]$LogDir,

        [Parameter(Mandatory = $false)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $false)]
        [string]$AppSettingsPath,

        [Parameter(Mandatory = $false)]
        [switch]$UseAppSettingsQueueRoot
    )

    if (-not $RepoRoot) {
        $here = $PSScriptRoot
        if (-not $here) {
            $here = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        if ($here) {
            $RepoRoot = Split-Path -Parent (Split-Path -Parent $here)
        } else {
            $RepoRoot = (Get-Location).Path
        }
    }

    $localConfigPath = Join-Path $RepoRoot 'config\qc-debug-locations.local.json'
    $localCfg = $null
    if (Test-Path -LiteralPath $localConfigPath) {
        try {
            $localCfg = Get-Content -LiteralPath $localConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
        } catch {
            Write-Warning ("Failed to parse local debug locations config: {0}" -f $localConfigPath)
            $localCfg = $null
        }
    }

    $queueRootSource = $null
    $resolvedQueueRoot = $null

    if ($QueueRoot -and $QueueRoot.Trim().Length -gt 0) {
        $resolvedQueueRoot = $QueueRoot.Trim()
        $queueRootSource = 'param'
    } elseif ($UseAppSettingsQueueRoot) {
        if (-not $AppSettingsPath) { $AppSettingsPath = Join-Path $RepoRoot 'appsettings.json' }
        if (-not (Test-Path -LiteralPath $AppSettingsPath)) {
            throw "appsettings.json not found: $AppSettingsPath"
        }
        $cfg = Get-Content -LiteralPath $AppSettingsPath -Raw | ConvertFrom-Json
        if ($cfg.queue -and $cfg.queue.rootDir) { $resolvedQueueRoot = [string]$cfg.queue.rootDir }
        elseif ($cfg.queue -and $cfg.queue.root) { $resolvedQueueRoot = [string]$cfg.queue.root }
        else { $resolvedQueueRoot = Join-Path $RepoRoot 'queue' }
        $queueRootSource = 'appsettings'
    } elseif ($env:QC_DEBUG_QUEUE_ROOT -and $env:QC_DEBUG_QUEUE_ROOT.Trim().Length -gt 0) {
        $resolvedQueueRoot = $env:QC_DEBUG_QUEUE_ROOT.Trim()
        $queueRootSource = 'env'
    } elseif ($localCfg -and $localCfg.queueRoot -and ([string]$localCfg.queueRoot).Trim().Length -gt 0) {
        $resolvedQueueRoot = ([string]$localCfg.queueRoot).Trim()
        $queueRootSource = 'local-config'
    } else {
        $resolvedQueueRoot = $script:QCDebugDefaultQueueRoot
        $queueRootSource = 'default'
    }

    $logDirSource = $null
    $resolvedLogDir = $null

    if ($LogDir -and $LogDir.Trim().Length -gt 0) {
        $resolvedLogDir = $LogDir.Trim()
        $logDirSource = 'param'
    } elseif ($env:QC_DEBUG_LOG_DIR -and $env:QC_DEBUG_LOG_DIR.Trim().Length -gt 0) {
        $resolvedLogDir = $env:QC_DEBUG_LOG_DIR.Trim()
        $logDirSource = 'env'
    } elseif ($localCfg -and $localCfg.logDir -and ([string]$localCfg.logDir).Trim().Length -gt 0) {
        $resolvedLogDir = ([string]$localCfg.logDir).Trim()
        $logDirSource = 'local-config'
    } else {
        $resolvedLogDir = Join-Path $resolvedQueueRoot '_logs'
        $logDirSource = 'derived'
    }

    [pscustomobject]@{
        QueueRoot       = $resolvedQueueRoot
        LogDir          = $resolvedLogDir
        QueueRootSource = $queueRootSource
        LogDirSource    = $logDirSource
        LocalConfigPath = $localConfigPath
        RepoRoot        = $RepoRoot
    }
}

# Direct execution: resolve with CLI-style args (no script-level param block so
# dot-sourcing from Show-QCQueueDiag does not steal parent parameters).
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch '^\s*\.') {
    $directParams = @{}
    for ($i = 0; $i -lt $args.Count; $i++) {
        $a = [string]$args[$i]
        if ($a -eq '-QueueRoot' -and ($i + 1) -lt $args.Count) { $directParams['QueueRoot'] = [string]$args[++$i]; continue }
        if ($a -eq '-LogDir' -and ($i + 1) -lt $args.Count) { $directParams['LogDir'] = [string]$args[++$i]; continue }
        if ($a -eq '-RepoRoot' -and ($i + 1) -lt $args.Count) { $directParams['RepoRoot'] = [string]$args[++$i]; continue }
        if ($a -eq '-AppSettingsPath' -and ($i + 1) -lt $args.Count) { $directParams['AppSettingsPath'] = [string]$args[++$i]; continue }
        if ($a -eq '-UseAppSettingsQueueRoot') { $directParams['UseAppSettingsQueueRoot'] = $true; continue }
    }
    Resolve-QCDebugLocations @directParams
}
