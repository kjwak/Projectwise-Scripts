# Verifies appsettings.json with // and /* */ comments loads via Read-QCAppSettings.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_cfg_comment_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
$cfgPath = Join-Path $tempRoot 'appsettings.json'
@'
{
  // line comment
  "dryRun": true,
  "watcher": { /* block */ "mode": "audit_only" }
}
'@ | Set-Content -LiteralPath $cfgPath -Encoding UTF8

try {
    $res = Read-QCAppSettings -Path $cfgPath
    if (-not $res.IsSuccess) { throw $res.Message }
    if (-not [bool]$res.Data.config.dryRun) { throw 'dryRun should be true' }
    if ([string]$res.Data.config.watcher.mode -ne 'audit_only') { throw 'watcher.mode parse failed' }
    Write-Host 'OK: JSON comment stripping in Read-QCAppSettings' -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# Repo appsettings is strict JSON; optional // comments still supported by loader.
$repoCfg = Join-Path $repoRoot 'appsettings.json'
$repoRes = Read-QCAppSettings -Path $repoCfg
if (-not $repoRes.IsSuccess) { throw ('repo appsettings failed: ' + $repoRes.Message) }
if (-not $repoRes.Data.config.ContainsKey('projectWise')) { throw 'projectWise missing' }
Write-Host 'OK: repo appsettings.json loads' -ForegroundColor Green
