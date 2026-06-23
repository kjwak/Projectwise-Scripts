<#
.SYNOPSIS
Copies QC pipeline modules and entrypoint scripts to a worker/install root and optionally restarts the dashboard.

.DESCRIPTION
Use after merging lane-notification, DOCUMENT_DELETE registry, and reset-script changes to dev.
Copies modules/, email/ (notification HTML template and logo), scripts/Watch-QCTrigger.ps1,
scripts/Run-QCProcessor.ps1, scripts/Import-QCScriptModules.ps1, scripts/Start-QCPipelineDashboard.ps1,
and scripts/Reset-QCFolderWorkflow.ps1 to the target root.

Does not copy appsettings.json, appsettings.local.json, appsettings.secrets.json, or queue data.
Restart the dashboard after publish so loaded modules refresh.

.EXAMPLE
.\scripts\Publish-QCPipelineCode.ps1 -WorkerRoot 'D:\QC_Pipeline\Prepend PDF QC'
.\scripts\Publish-QCPipelineCode.ps1 -WorkerRoot 'D:\QC_Pipeline\Prepend PDF QC' -ConfirmRestart
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$WorkerRoot,
    [switch]$ConfirmRestart,
    [switch]$SkipModuleCopy
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$workerRootResolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkerRoot)
if (-not (Test-Path -LiteralPath $workerRootResolved)) {
    throw "WorkerRoot not found: $workerRootResolved"
}

$copyPlan = @(
    @{ src = (Join-Path $repoRoot 'modules'); dst = (Join-Path $workerRootResolved 'modules'); type = 'dir' }
    @{ src = (Join-Path $repoRoot 'email'); dst = (Join-Path $workerRootResolved 'email'); type = 'dir' }
    @{ src = (Join-Path $repoRoot 'scripts\Watch-QCTrigger.ps1'); dst = (Join-Path $workerRootResolved 'scripts\Watch-QCTrigger.ps1'); type = 'file' }
    @{ src = (Join-Path $repoRoot 'scripts\Run-QCProcessor.ps1'); dst = (Join-Path $workerRootResolved 'scripts\Run-QCProcessor.ps1'); type = 'file' }
    @{ src = (Join-Path $repoRoot 'scripts\Import-QCScriptModules.ps1'); dst = (Join-Path $workerRootResolved 'scripts\Import-QCScriptModules.ps1'); type = 'file' }
    @{ src = (Join-Path $repoRoot 'scripts\Start-QCPipelineDashboard.ps1'); dst = (Join-Path $workerRootResolved 'scripts\Start-QCPipelineDashboard.ps1'); type = 'file' }
    @{ src = (Join-Path $repoRoot 'scripts\Reset-QCFolderWorkflow.ps1'); dst = (Join-Path $workerRootResolved 'scripts\Reset-QCFolderWorkflow.ps1'); type = 'file' }
)

Write-Host "[Publish] Repo: $repoRoot" -ForegroundColor Cyan
Write-Host "[Publish] Worker: $workerRootResolved" -ForegroundColor Cyan

if (-not $SkipModuleCopy.IsPresent) {
    foreach ($item in $copyPlan) {
        if (-not (Test-Path -LiteralPath $item.src)) {
            throw "Missing source: $($item.src)"
        }
        if ($item.type -eq 'dir') {
            if ($PSCmdlet.ShouldProcess($item.dst, 'Mirror modules directory')) {
                if (-not (Test-Path -LiteralPath $item.dst)) {
                    New-Item -ItemType Directory -Path $item.dst -Force | Out-Null
                }
                Copy-Item -Path (Join-Path $item.src '*') -Destination $item.dst -Recurse -Force
                Write-Host "  Copied modules -> $($item.dst)" -ForegroundColor Green
            }
        } else {
            $dstDir = Split-Path -Parent $item.dst
            if (-not (Test-Path -LiteralPath $dstDir)) {
                New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            }
            if ($PSCmdlet.ShouldProcess($item.dst, 'Copy script')) {
                Copy-Item -Path $item.src -Destination $item.dst -Force
                Write-Host "  Copied $(Split-Path -Leaf $item.src)" -ForegroundColor Green
            }
        }
    }
}

if ($ConfirmRestart.IsPresent) {
    $stopScript = Join-Path $workerRootResolved 'scripts\Stop-QCPipeline.ps1'
    $startScript = Join-Path $workerRootResolved 'scripts\Start-QCPipelineDashboard.ps1'
    if (-not (Test-Path -LiteralPath $stopScript)) {
        Write-Host '[Publish] Stop-QCPipeline.ps1 not found; restart manually.' -ForegroundColor Yellow
    } elseif ($PSCmdlet.ShouldProcess('QC pipeline', 'Restart dashboard')) {
        & $stopScript
        Start-Sleep -Seconds 3
        if (Test-Path -LiteralPath $startScript) {
            Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $startScript) -WorkingDirectory $workerRootResolved
            Write-Host '[Publish] Dashboard restart launched.' -ForegroundColor Green
        }
    }
} else {
    Write-Host '[Publish] Done. Restart Watch-QCTrigger / dashboard to load new modules (or pass -ConfirmRestart).' -ForegroundColor Yellow
}
