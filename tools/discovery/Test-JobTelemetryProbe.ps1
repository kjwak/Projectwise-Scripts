<#
.SYNOPSIS
Probes Write-QCJobTelemetry against processing_jobs (connectivity + MERGE + read-back).

.PARAMETER AppSettingsPath
Path to appsettings.json. Defaults to repo root.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = ''
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

$modulesDir = Join-Path $repoRoot 'modules'
foreach ($mod in @('Core.Results.psm1', 'Core.Database.psm1')) {
    $modPath = Join-Path $modulesDir $mod
    if (-not (Test-Path -LiteralPath $modPath)) {
        throw "Module not found: $modPath"
    }
    Import-Module $modPath -Force -ErrorAction Stop
}

if (-not (Test-Path -LiteralPath $AppSettingsPath)) {
    throw "appsettings.json not found: $AppSettingsPath"
}

function _DeepHashtable ($obj) {
    if ($obj -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = _DeepHashtable $p.Value }
        return $h
    }
    if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
        return @($obj | ForEach-Object { _DeepHashtable $_ })
    }
    return $obj
}

Write-Host "Loading config from: $AppSettingsPath"
$raw = Get-Content -LiteralPath $AppSettingsPath -Raw -ErrorAction Stop
$config = _DeepHashtable ($raw | ConvertFrom-Json -ErrorAction Stop)

Write-Host "DB enabled: $(Test-QCDatabaseEnabled -Config $config)"

$init = Initialize-QCDatabaseSchema -Config $config
Write-Host "Schema: $($init.Code) $($init.Message)"

$probeId = 'test_telemetry_probe'
$tel = Write-QCJobTelemetry -Config $config -JobId $probeId -JobType 'STATUS_SET_GEN' -Status 'succeeded' -SourceFolder 'test' -DurationMs 1
Write-Host "Telemetry: $($tel.Code) $($tel.Message) success=$($tel.IsSuccess)"

$sql = @"
SELECT job_id, job_type, status, completed_at
FROM processing_jobs
WHERE job_id IN (@probeId, @realId)
"@
$q = Invoke-QCDatabaseQuery -Config $config -Sql $sql -Parameters @{
    probeId = $probeId
    realId  = 'qc_statussetgen_fec4de30b10547f9'
}
if (-not $q.IsSuccess) { throw $q.Message }
$q.Data.table | Format-Table -AutoSize
