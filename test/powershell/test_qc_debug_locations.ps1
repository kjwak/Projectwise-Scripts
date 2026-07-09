<#
.SYNOPSIS
Smoke tests for Resolve-QCDebugLocations precedence and Show-QCQueueDiag wiring.

Does not mutate production appsettings.json or the live queue share.
#>

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)

. (Join-Path $repoRoot 'scripts\diagnostics\Resolve-QCDebugLocations.ps1')

$failed = 0
function Assert-Eq {
    param([string]$Name, $Expected, $Actual)
    if ("$Actual" -ne "$Expected") {
        Write-Host ("FAIL: {0} expected={1} actual={2}" -f $Name, $Expected, $Actual) -ForegroundColor Red
        $script:failed++
    } else {
        Write-Host ("PASS: {0}" -f $Name) -ForegroundColor Green
    }
}

$tempRoot = Join-Path $env:TEMP ("qc-debug-loc-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'config') -Force | Out-Null
$localCfgPath = Join-Path $tempRoot 'config\qc-debug-locations.local.json'
@{
    queueRoot = 'C:\fake-from-local-config\queue'
    logDir    = 'C:\fake-from-local-config\logs'
} | ConvertTo-Json | Set-Content -LiteralPath $localCfgPath -Encoding UTF8

# Isolate env for tests
$prevQueueEnv = $env:QC_DEBUG_QUEUE_ROOT
$prevLogEnv = $env:QC_DEBUG_LOG_DIR
try {
    Remove-Item Env:QC_DEBUG_QUEUE_ROOT -ErrorAction SilentlyContinue
    Remove-Item Env:QC_DEBUG_LOG_DIR -ErrorAction SilentlyContinue

    # local config beats default
    $r = Resolve-QCDebugLocations -RepoRoot $tempRoot
    Assert-Eq 'local-config queueRoot' 'C:\fake-from-local-config\queue' $r.QueueRoot
    Assert-Eq 'local-config QueueRootSource' 'local-config' $r.QueueRootSource
    Assert-Eq 'local-config logDir' 'C:\fake-from-local-config\logs' $r.LogDir
    Assert-Eq 'local-config LogDirSource' 'local-config' $r.LogDirSource

    # env beats local config
    $env:QC_DEBUG_QUEUE_ROOT = 'C:\fake-from-env\queue'
    $env:QC_DEBUG_LOG_DIR = 'C:\fake-from-env\logs'
    $r = Resolve-QCDebugLocations -RepoRoot $tempRoot
    Assert-Eq 'env queueRoot' 'C:\fake-from-env\queue' $r.QueueRoot
    Assert-Eq 'env QueueRootSource' 'env' $r.QueueRootSource
    Assert-Eq 'env logDir' 'C:\fake-from-env\logs' $r.LogDir
    Assert-Eq 'env LogDirSource' 'env' $r.LogDirSource

    # param beats env
    $r = Resolve-QCDebugLocations -RepoRoot $tempRoot -QueueRoot 'C:\fake-from-param\queue' -LogDir 'C:\fake-from-param\logs'
    Assert-Eq 'param queueRoot' 'C:\fake-from-param\queue' $r.QueueRoot
    Assert-Eq 'param QueueRootSource' 'param' $r.QueueRootSource
    Assert-Eq 'param logDir' 'C:\fake-from-param\logs' $r.LogDir
    Assert-Eq 'param LogDirSource' 'param' $r.LogDirSource

    # derived log dir when queue root set and no logDir from env/local config
    Remove-Item Env:QC_DEBUG_QUEUE_ROOT -ErrorAction SilentlyContinue
    Remove-Item Env:QC_DEBUG_LOG_DIR -ErrorAction SilentlyContinue
    $emptyRoot = Join-Path $env:TEMP ("qc-debug-empty-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path (Join-Path $emptyRoot 'config') -Force | Out-Null
    $r = Resolve-QCDebugLocations -RepoRoot $emptyRoot -QueueRoot 'C:\fake-derived\queue'
    Assert-Eq 'derived LogDir' (Join-Path 'C:\fake-derived\queue' '_logs') $r.LogDir
    Assert-Eq 'derived LogDirSource' 'derived' $r.LogDirSource

    # default when no local config / env / param
    $r = Resolve-QCDebugLocations -RepoRoot $emptyRoot
    Assert-Eq 'default queueRoot' '\\192.168.22.90\QC_Queue' $r.QueueRoot
    Assert-Eq 'default QueueRootSource' 'default' $r.QueueRootSource
    Assert-Eq 'default derived LogDir' (Join-Path '\\192.168.22.90\QC_Queue' '_logs') $r.LogDir
    Remove-Item -LiteralPath $emptyRoot -Recurse -Force -ErrorAction SilentlyContinue
}
finally {
    if ($null -ne $prevQueueEnv) { $env:QC_DEBUG_QUEUE_ROOT = $prevQueueEnv } else { Remove-Item Env:QC_DEBUG_QUEUE_ROOT -ErrorAction SilentlyContinue }
    if ($null -ne $prevLogEnv) { $env:QC_DEBUG_LOG_DIR = $prevLogEnv } else { Remove-Item Env:QC_DEBUG_LOG_DIR -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# Unreachable queue root fails clearly (path + source)
$unreachable = 'C:\qc-debug-locations-unreachable-' + [guid]::NewGuid().ToString('N')
$diagScript = Join-Path $repoRoot 'scripts\diagnostics\Show-QCQueueDiag.ps1'
$errText = ''
try {
    & $diagScript -QueueRoot $unreachable -NoLogs 2>&1 | Out-Null
    Write-Host 'FAIL: unreachable queue root should throw' -ForegroundColor Red
    $failed++
} catch {
    $errText = "$_"
    if ($errText -match [regex]::Escape($unreachable) -and $errText -match 'source=param') {
        Write-Host 'PASS: unreachable queue root reports path + source=param' -ForegroundColor Green
    } else {
        Write-Host ("FAIL: unreachable error missing path/source: {0}" -f $errText) -ForegroundColor Red
        $failed++
    }
}

# appsettings.json queue.rootDir unchanged and not mutated by resolver
$appsettingsPath = Join-Path $repoRoot 'appsettings.json'
$beforeHash = (Get-FileHash -LiteralPath $appsettingsPath -Algorithm SHA256).Hash
$cfg = Get-Content -LiteralPath $appsettingsPath -Raw | ConvertFrom-Json
$prodRoot = [string]$cfg.queue.rootDir
Assert-Eq 'appsettings queue.rootDir value' 'C:\QC_E2E_RealRun\queue' $prodRoot

# Resolver without -UseAppSettingsQueueRoot must NOT return production queue.rootDir
Remove-Item Env:QC_DEBUG_QUEUE_ROOT -ErrorAction SilentlyContinue
$r = Resolve-QCDebugLocations -RepoRoot $repoRoot
if ($r.QueueRoot -eq $prodRoot) {
    Write-Host ("FAIL: debug resolver returned production queue.rootDir unexpectedly: {0} source={1}" -f $r.QueueRoot, $r.QueueRootSource) -ForegroundColor Red
    $failed++
} else {
    Write-Host ("PASS: debug resolver QueueRoot is not production appsettings path (got {0} source={1})" -f $r.QueueRoot, $r.QueueRootSource) -ForegroundColor Green
}

$afterHash = (Get-FileHash -LiteralPath $appsettingsPath -Algorithm SHA256).Hash
Assert-Eq 'appsettings.json not mutated' $beforeHash $afterHash

# Confirm local.json is gitignored (git check-ignore)
Push-Location $repoRoot
try {
    $ignored = git check-ignore -v 'config/qc-debug-locations.local.json' 2>$null
    if ($LASTEXITCODE -eq 0 -and $ignored) {
        Write-Host 'PASS: config/qc-debug-locations.local.json is gitignored' -ForegroundColor Green
    } else {
        Write-Host 'FAIL: config/qc-debug-locations.local.json is not gitignored' -ForegroundColor Red
        $failed++
    }
} finally {
    Pop-Location
}

if ($failed -gt 0) {
    Write-Host ("`n{0} assertion(s) failed." -f $failed) -ForegroundColor Red
    exit 1
}
Write-Host "`nAll smoke tests passed." -ForegroundColor Green
exit 0
