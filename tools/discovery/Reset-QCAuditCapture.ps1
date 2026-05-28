<#
.SYNOPSIS
Reset audit-trail capture so the next watcher tick uses initialLookbackSeconds.

.DESCRIPTION
Removes the local watermark file. Optionally clears poll_runs.watermark_after in SQL
(use when you need a full re-capture without deleting the whole queue folder).

Does not truncate audit_events.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [switch]$ClearPollRunWatermarks
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$queueRoot = Join-Path $repoRoot 'queue'
if ($config.queue -and $config.queue.rootDir) { $queueRoot = [string]$config.queue.rootDir }
$wmPath = Join-Path (Join-Path $queueRoot '_watcher') 'audit-capture-watermark.txt'

if (Test-Path -LiteralPath $wmPath) {
    Remove-Item -LiteralPath $wmPath -Force
    Write-Host "Removed watermark file: $wmPath" -ForegroundColor Green
} else {
    Write-Host "No watermark file at: $wmPath" -ForegroundColor DarkYellow
}

if ($ClearPollRunWatermarks.IsPresent) {
    if (-not (Test-QCDatabaseEnabled -Config $config)) {
        Write-Host 'SKIP: database.enabled is false; cannot clear poll_runs watermarks.' -ForegroundColor Yellow
    } else {
        $res = Invoke-QCDatabaseNonQuery -Config $config -Sql 'UPDATE poll_runs SET watermark_after = NULL WHERE watermark_after IS NOT NULL'
        if ($res.IsSuccess) {
            Write-Host "Cleared poll_runs.watermark_after ($([int]$res.Data.rowsAffected) row(s))." -ForegroundColor Green
        } else {
            throw $res.Message
        }
    }
}

Write-Host 'Next watcher audit tick will use initialLookbackSeconds when the watermark file is absent.' -ForegroundColor Cyan
