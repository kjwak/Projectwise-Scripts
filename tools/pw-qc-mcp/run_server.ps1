# Fast launcher: copy server.ps1 to LOCALAPPDATA (avoids OneDrive latency), then run in-process.
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$cacheDir = Join-Path $env:LOCALAPPDATA 'pw-qc-mcp'
$src = Join-Path $here 'server.ps1'
$dst = Join-Path $cacheDir 'server.ps1'
New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
if (-not (Test-Path -LiteralPath $dst) -or (Get-Item -LiteralPath $src).LastWriteTimeUtc -gt (Get-Item -LiteralPath $dst).LastWriteTimeUtc) {
    Copy-Item -LiteralPath $src -Destination $dst -Force
}
if ([string]::IsNullOrWhiteSpace($env:PWQC_REPO_ROOT)) {
    $env:PWQC_REPO_ROOT = (Resolve-Path (Join-Path $here '..\..')).Path
}
. $dst
