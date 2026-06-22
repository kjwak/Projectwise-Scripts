<#
.SYNOPSIS
Backfills automation_events from historical Run-QCProcessor_*.jsonl and Watch-QCTrigger_*.jsonl files.

.DESCRIPTION
Parses JSONL safely, applies the same heartbeat filters as Write-QCAutomationEvent,
and uses dedupe_key for idempotent re-runs.

.EXAMPLE
.\scripts\Import-QCJsonlLogsToAutomationEvents.ps1 -LogDirectory C:\PW_QC_LOCAL\queue\_logs
.\scripts\Import-QCJsonlLogsToAutomationEvents.ps1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$AppSettingsPath = '',
    [string]$LogDirectory = '',
    [string[]]$FilePatterns = @('Run-QCProcessor_*.jsonl', 'Watch-QCTrigger_*.jsonl'),
    [int]$MaxFiles = 0
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Database\Core.Database.psm1'
    'Core\Core.Telemetry.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Test-QCDatabaseEnabled'
    'Get-QCAutomationTelemetrySettings'
    'Write-QCAutomationEvent'
) -Context 'Import-QCJsonlLogsToAutomationEvents bootstrap'

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled is false; cannot import automation_events.'
}

$schemaRes = Initialize-QCDatabaseSchema -Config $config
if (-not $schemaRes.IsSuccess) { throw "Schema init failed: $($schemaRes.Message)" }

if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
    if ($env:QC_JSON_LOG_DIR) { $LogDirectory = [string]$env:QC_JSON_LOG_DIR }
    else {
        $settings = Get-QCAutomationTelemetrySettings -Config $config
        if ($settings.jsonLogDir) { $LogDirectory = [string]$settings.jsonLogDir }
    }
}
if ([string]::IsNullOrWhiteSpace($LogDirectory) -or -not (Test-Path -LiteralPath $LogDirectory)) {
    throw "Log directory not found. Pass -LogDirectory or set QC_JSON_LOG_DIR."
}

$files = @()
foreach ($pat in $FilePatterns) {
    $files += @(Get-ChildItem -LiteralPath $LogDirectory -Filter $pat -File -ErrorAction SilentlyContinue)
}
$files = @($files | Sort-Object LastWriteTime | Select-Object -Unique FullName)
if ($MaxFiles -gt 0) { $files = @($files | Select-Object -Last $MaxFiles) }

$stats = @{
    files = $files.Count
    lines = 0
    inserted = 0
    skipped_filter = 0
    skipped_duplicate = 0
    parse_errors = 0
    write_errors = 0
}

foreach ($f in $files) {
    $processName = if ($f.Name -like 'Run-QCProcessor_*') { 'Run-QCProcessor' }
        elseif ($f.Name -like 'Watch-QCTrigger_*') { 'Watch-QCTrigger' }
        else { 'unknown' }

    Write-Host "Importing $($f.Name) as $processName ..." -ForegroundColor Cyan
    $reader = $null
    try {
        $reader = [System.IO.StreamReader]::new($f.FullName)
        while ($null -ne ($line = $reader.ReadLine())) {
            $stats.lines++
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($PSCmdlet.ShouldProcess($f.Name, "import line $($stats.lines)")) {
                $res = Import-QCAutomationEventFromJsonLine -Line $line -ProcessName $processName -Config $config
                if ($res.Data.inserted) { $stats.inserted++ }
                elseif ($res.Data.duplicate) { $stats.skipped_duplicate++ }
                elseif ($res.Data.skipped) { $stats.skipped_filter++ }
                elseif ($res.Code -eq 'IMPORT_PARSE_FAILED') { $stats.parse_errors++ }
                elseif (-not $res.IsSuccess) { $stats.write_errors++ }
            }
        }
    } finally {
        if ($reader) { $reader.Dispose() }
    }
}

Write-Host ''
Write-Host 'Import summary:' -ForegroundColor Green
$stats.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("  {0}: {1}" -f $_.Key, $_.Value) }
