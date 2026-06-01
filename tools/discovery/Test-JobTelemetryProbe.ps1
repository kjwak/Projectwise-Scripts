$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

$cfgPath = Join-Path $repoRoot 'appsettings.json'
$config = [hashtable](Read-QCAppSettings -Path $cfgPath).Data.config

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
