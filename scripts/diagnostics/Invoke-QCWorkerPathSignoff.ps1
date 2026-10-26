<#
.SYNOPSIS
Run Phase 4 server path audit checklist (docs/archive/phase/phase-4-server-path-audit.md §9) on a worker host.

.DESCRIPTION
Read-only inspection: running QC processes, Task Scheduler actions, appsettings path overrides,
publish-vs-clone drift. ProjectWise Rules Engine (§9.4) must still be checked manually in PW Administrator.

.EXAMPLE
.\scripts\diagnostics\Invoke-QCWorkerPathSignoff.ps1 -WorkerRoot 'D:\QC_Pipeline\Prepend PDF QC'
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$WorkerRoot = 'D:\QC_Pipeline\Prepend PDF QC'
)

$ErrorActionPreference = 'Stop'
$hostName = $env:COMPUTERNAME
$workerResolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkerRoot)

Write-Host "=== QC worker path sign-off ($hostName) ===" -ForegroundColor Cyan
Write-Host "Worker root: $workerResolved"
Write-Host ''

if (-not (Test-Path -LiteralPath $workerResolved)) {
    Write-Host "Worker root not found on this host. Run this script on the production worker (or pass -WorkerRoot)." -ForegroundColor Yellow
    exit 2
}

Write-Host '--- 9.1 Running QC pipeline processes ---' -ForegroundColor Cyan
$patterns = 'Start-QCPipelineDashboard|Start-QCRemoteWorkerHost|Watch-QCTrigger|Run-QCProcessor'
$procs = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -and ($_.CommandLine -match $patterns) })
if ($procs.Count -eq 0) {
    Write-Host '  (none)'
    $processSummary = 'none running'
} else {
    $processSummary = @($procs | ForEach-Object {
        $cmd = $_.CommandLine
        if ($cmd.Length -gt 180) { $cmd = $cmd.Substring(0, 180) + '...' }
        "pid=$($_.ProcessId) $cmd"
    }) -join '; '
    foreach ($p in $procs) {
        $cmd = if ($p.CommandLine.Length -gt 200) { $p.CommandLine.Substring(0, 200) + '...' } else { $p.CommandLine }
        Write-Host "  pid=$($p.ProcessId) $cmd"
    }
}

Write-Host ''
Write-Host '--- 9.2 Task Scheduler (QC-related) ---' -ForegroundColor Cyan
$taskSummary = 'none'
try {
    $tasks = @(Get-ScheduledTask -ErrorAction Stop |
        Where-Object { $_.Actions.Execute -match 'powershell|pwsh' } |
        ForEach-Object {
            $a = $_.Actions
            [PSCustomObject]@{ Task = $_.TaskName; Execute = $a.Execute; Arguments = $a.Arguments }
        } |
        Where-Object { $_.Arguments -match 'QC|Prepend|Watch-QCTrigger|Run-QCProcessor|Start-QCPipeline|Start-QCRemoteWorkerHost' })
    if ($tasks.Count -eq 0) {
        Write-Host '  (no matching scheduled tasks)'
    } else {
        $taskSummary = @($tasks | ForEach-Object { "$($_.Task): $($_.Arguments)" }) -join '; '
        $tasks | Format-Table -AutoSize | Out-String | Write-Host
    }
} catch {
    Write-Host "  Task Scheduler query failed: $($_.Exception.Message)" -ForegroundColor Yellow
    $taskSummary = "query failed: $($_.Exception.Message)"
}

Write-Host '--- 9.3 appsettings path overrides ---' -ForegroundColor Cyan
$overrideKeys = @(
    'qcPrepend.legacyScriptPath'
    'qcPrepend.projectWiseScriptPath'
    'qcPrepend.overlayExePath'
    'statusSet.legacyScriptPath'
    'notifications.email.templatePath'
)
$configSummary = 'none (no appsettings*.json)'
$configFiles = @(Get-ChildItem -LiteralPath $workerResolved -Filter 'appsettings*.json' -File -ErrorAction SilentlyContinue)
if ($configFiles.Count -gt 0) {
    Write-Host "  Config files: $(@($configFiles.Name) -join ', ')"
    $foundOverrides = @()
    foreach ($cf in $configFiles) {
        if ($cf.Name -match 'secrets') { continue }
        try {
            $json = Get-Content -LiteralPath $cf.FullName -Raw | ConvertFrom-Json
            foreach ($key in $overrideKeys) {
                $parts = $key.Split('.')
                $node = $json
                foreach ($part in $parts) {
                    if ($null -eq $node) { break }
                    $node = $node.$part
                }
                if ($node -and [string]$node.Trim()) {
                    $foundOverrides += "$($cf.Name): $key=$node"
                    Write-Host "  $($cf.Name) -> $key = $node"
                }
            }
        } catch {
            Write-Host "  Could not parse $($cf.Name): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    if ($foundOverrides.Count -gt 0) {
        $configSummary = $foundOverrides -join '; '
    } else {
        $configSummary = 'no script-path overrides in appsettings*.json (excluding secrets)'
    }
} else {
    Write-Host '  (no appsettings*.json under worker root)'
}

Write-Host ''
Write-Host '--- 9.4 ProjectWise Rules Engine ---' -ForegroundColor Cyan
Write-Host '  Manual: search PW Rules Engine for prepend_qc, combine_status_set, Watch-QCTrigger, Run-QCProcessor, Start-QCPipeline.' -ForegroundColor Yellow
$rulesSummary = 'manual check required'

Write-Host ''
Write-Host '--- 9.5 Publish vs clone drift ---' -ForegroundColor Cyan
$driftPaths = @(
    'email\templates\qc_notification.html'
    'scripts\Restore-QCModuleExports.ps1'
    'scripts\service\Stop-QCPipeline.ps1'
    'scripts\service\Start-QCPipelineDashboard.ps1'
    'scripts\maintenance\Reset-QCFolderWorkflow.ps1'
    'legacy\prepend_qc.ps1'
    'dist\qc_overlay_prepend\qc_overlay_prepend.exe'
)
$drift = foreach ($rel in $driftPaths) {
    $p = Join-Path $workerResolved $rel
    [PSCustomObject]@{ Path = $rel; Exists = (Test-Path -LiteralPath $p) }
}
$drift | Format-Table -AutoSize | Out-String | Write-Host

Write-Host '--- 10 Sign-off row (paste into docs/archive/phase/phase-4-server-path-audit.md) ---' -ForegroundColor Cyan
$date = Get-Date -Format 'yyyy-MM-dd'
$row = "| $hostName | ``$workerResolved`` | $taskSummary | $rulesSummary | $configSummary | $processSummary | | $date |"
Write-Host $row
Write-Host ''
Write-Host 'Sign-off complete only after Rules Engine manual check and reviewer name filled in.' -ForegroundColor Yellow
