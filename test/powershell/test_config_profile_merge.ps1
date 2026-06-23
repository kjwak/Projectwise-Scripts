# Verifies appsettings profile merge (base + overlay + local).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force

function _Assert($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('qc_cfg_merge_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

$basePath = Join-Path $tempRoot 'appsettings.json'
$profilePath = Join-Path $tempRoot 'appsettings.test.json'
$localPath = Join-Path $tempRoot 'appsettings.test.local.json'

@'
{
  "dryRun": false,
  "database": {
    "enabled": true,
    "connectionString": "Server=base;Database=BaseDb;"
  },
  "watcher": { "mode": "hybrid" }
}
'@ | Set-Content -LiteralPath $basePath -Encoding UTF8

@'
{
  "dryRun": true,
  "database": {
    "connectionString": "Server=test;Database=TestDb;",
    "allowWritesInDryRun": true
  }
}
'@ | Set-Content -LiteralPath $profilePath -Encoding UTF8

@'
{
  "database": {
    "connectionString": "Server=local;Database=LocalDb;"
  }
}
'@ | Set-Content -LiteralPath $localPath -Encoding UTF8

try {
    $chain = @(Resolve-QCAppSettingsMergeChain -Path $profilePath)
    _Assert ($chain.Count -eq 3) ('expected 3 merge files, got ' + $chain.Count)
    _Assert ($chain[0] -like '*appsettings.json') 'first file should be base appsettings.json'
    _Assert ($chain[-1] -like '*appsettings.test.local.json') 'last file should be profile local'

    $res = Read-QCAppSettings -Path $profilePath
    _Assert $res.IsSuccess $res.Message
    $cfg = $res.Data.config
    _Assert ([bool]$cfg.dryRun) 'profile should set dryRun true'
    _Assert ($cfg.watcher.mode -eq 'hybrid') 'base watcher.mode should remain'
    _Assert ($cfg.database.enabled -eq $true) 'database.enabled should remain from base'
    _Assert ($cfg.database.allowWritesInDryRun -eq $true) 'profile should merge database.allowWritesInDryRun'
    _Assert ($cfg.database.connectionString -like '*LocalDb*') 'local overlay should win connectionString'

    $localOnly = Join-Path $tempRoot 'appsettings.local.json'
    '{ "dryRun": false }' | Set-Content -LiteralPath $localOnly -Encoding UTF8
    $baseRes = Read-QCAppSettings -Path $basePath
    _Assert $baseRes.IsSuccess $baseRes.Message
    _Assert (-not [bool]$baseRes.Data.config.dryRun) 'appsettings.local.json should override base dryRun'

    $secretsPath = Join-Path $tempRoot 'appsettings.secrets.json'
    @'
{
  "notifications": {
    "graph": { "clientSecret": "from-secrets-file" }
  }
}
'@ | Set-Content -LiteralPath $secretsPath -Encoding UTF8
    $secretsRes = Read-QCAppSettings -Path $basePath
    _Assert $secretsRes.IsSuccess $secretsRes.Message
    _Assert ($secretsRes.Data.config.notifications.graph.clientSecret -eq 'from-secrets-file') 'appsettings.secrets.json should merge graph secrets'

    Write-Host 'OK: config profile merge' -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
