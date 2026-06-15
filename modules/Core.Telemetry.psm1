# Core.Telemetry.psm1
# Durable automation_events telemetry + JSONL event routing helpers.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
if (-not (Get-Command -Name 'Invoke-QCDatabaseNonQuery' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'Core.Database.psm1') -Force
}

$script:QCT_Config = $null
$script:QCT_ProcessName = ''
$script:QCT_RunId = ''

$script:QCT_DefaultExcludeCodes = @(
    'WATCH_TICK_START'
    'WATCH_TICK_SLEEP'
    'WORKER_NO_JOB'
)

$script:QCT_WorkerStageNoisePatterns = @(
    '(?i)polling queue'
    '(?i)queue poll'
    '(?i)idle sleep'
    '(?i)waiting for job'
)

function Set-QCAutomationTelemetryContext {
    <#
    .SYNOPSIS
    Binds appsettings and process identity for Write-QCAutomationEvent (call once at process start).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ProcessName,
        [string]$RunId = ''
    )
    $script:QCT_Config = $Config
    $script:QCT_ProcessName = ([string]$ProcessName).Trim()
    if ([string]::IsNullOrWhiteSpace($RunId)) {
        $RunId = [guid]::NewGuid().ToString('n')
    }
    $script:QCT_RunId = $RunId
    $settings = Get-QCAutomationTelemetrySettings -Config $Config
    if (Get-Command -Name 'Set-QCJsonLogRetentionSettings' -ErrorAction SilentlyContinue) {
        Set-QCJsonLogRetentionSettings -RetentionDays $settings.jsonLogRetentionDays -MaxFileSizeMb $settings.jsonLogMaxFileSizeMb -Config $Config
    }
    return $RunId
}

function Get-QCAutomationTelemetrySettings {
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $enabled = $true
    $retentionDays = 7
    $maxFileMb = 50
    $jsonLogDir = ''
    $exclude = @($script:QCT_DefaultExcludeCodes)

    try {
        $t = $Config['telemetry']
        if ($t -and $t.automationEvents) {
            $ae = $t.automationEvents
            if ($null -ne $ae.enabled) { $enabled = [bool]$ae.enabled }
            if ($ae.jsonLogRetentionDays) { $retentionDays = [int]$ae.jsonLogRetentionDays }
            if ($ae.jsonLogMaxFileSizeMb) { $maxFileMb = [int]$ae.jsonLogMaxFileSizeMb }
            if ($ae.jsonLogDir) { $jsonLogDir = [string]$ae.jsonLogDir }
            if ($ae.excludeCodes) { $exclude = @($ae.excludeCodes | ForEach-Object { [string]$_ }) }
        }
    } catch { }

    return @{
        enabled = $enabled
        excludeCodes = @($exclude | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        jsonLogRetentionDays = [Math]::Max(1, $retentionDays)
        jsonLogMaxFileSizeMb = [Math]::Max(1, $maxFileMb)
        jsonLogDir = $jsonLogDir
    }
}

function _QCT-FirstNonEmpty {
    param([hashtable]$Data, [string[]]$Keys)
    if (-not $Data) { return $null }
    foreach ($k in $Keys) {
        if (-not $Data.ContainsKey($k)) { continue }
        $v = $Data[$k]
        if ($null -eq $v -or $v -is [DBNull]) { continue }
        $s = [string]$v
        if (-not [string]::IsNullOrWhiteSpace($s)) { return $s.Trim() }
    }
    return $null
}

function Test-QCAutomationEventPersist {
    <#
    .SYNOPSIS
    Returns $true when an event should be written to automation_events.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Level,
        [Parameter(Mandatory)][string]$Code,
        [string]$Message = '',
        [hashtable]$Data = @{},
        [hashtable]$Config = @{}
    )

    $settings = Get-QCAutomationTelemetrySettings -Config $Config
    if (-not $settings.enabled) { return $false }

    $lvl = ([string]$Level).Trim()
    if ($lvl -match '^(?i)(warning|error)$') { return $true }

    $c = ([string]$Code).Trim()
    foreach ($ex in $settings.excludeCodes) {
        if ($c -ieq $ex) { return $false }
    }

    if ($c -ieq 'WORKER_STAGE') {
        $jobId = _QCT-FirstNonEmpty -Data $Data -Keys @('jobId', 'job_id')
        $docGuid = _QCT-FirstNonEmpty -Data $Data -Keys @('documentGuid', 'document_guid', 'docGuid', 'pw_objguid')
        $pkgId = _QCT-FirstNonEmpty -Data $Data -Keys @('sheetPackageId', 'sheet_package_id', 'packageId')
        if ($jobId -or $docGuid -or $pkgId) { return $true }

        $stageText = _QCT-FirstNonEmpty -Data $Data -Keys @('stage')
        if ([string]::IsNullOrWhiteSpace($stageText)) { $stageText = $Message }
        foreach ($pat in $script:QCT_WorkerStageNoisePatterns) {
            if ($stageText -match $pat) { return $false }
        }
        return $true
    }

    return $true
}

function New-QCAutomationEventDedupeKey {
    param(
        [string]$Timestamp,
        [string]$ProcessName,
        [string]$Code,
        [string]$Message,
        [string]$JobId,
        [string]$RunId
    )
    $parts = @(
        ([string]$Timestamp).Trim()
        ([string]$ProcessName).Trim().ToLowerInvariant()
        ([string]$Code).Trim().ToUpperInvariant()
        ([string]$Message).Trim()
        ([string]$JobId).Trim().ToLowerInvariant()
        ([string]$RunId).Trim().ToLowerInvariant()
    )
    $raw = ($parts -join '|')
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
        return ([BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    } catch {
        if ($raw.Length -gt 500) { $raw = $raw.Substring(0, 500) }
        return $raw
    }
}

function Write-QCAutomationEvent {
    <#
    .SYNOPSIS
    Inserts one automation_events row. Never throws; returns QCResult.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Level,
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Data = @{},
        [string]$Timestamp = '',
        [hashtable]$Config = $null,
        [string]$ProcessName = '',
        [string]$RunId = '',
        [string]$DedupeKey = '',
        [switch]$Force
    )

    if (-not $Data) { $Data = @{} }
    $cfg = if ($Config) { $Config } elseif ($script:QCT_Config) { $script:QCT_Config } else { @{} }
    $proc = if ($ProcessName) { $ProcessName } else { $script:QCT_ProcessName }
    if ([string]::IsNullOrWhiteSpace($proc)) { $proc = 'unknown' }
    $run = if ($RunId) { $RunId } else { $script:QCT_RunId }

    if (-not $Force -and -not (Test-QCAutomationEventPersist -Level $Level -Code $Code -Message $Message -Data $Data -Config $cfg)) {
        return New-QCSuccessResult -Code 'AUTOMATION_EVENT_SKIPPED' -Message 'Event filtered (heartbeat/noise).' -Data @{ written = $false; skipped = $true }
    }

    if (-not (Test-QCDatabaseEnabled -Config $cfg)) {
        return New-QCSuccessResult -Code 'AUTOMATION_EVENT_SKIPPED' -Message 'Database disabled.' -Data @{ written = $false; skipped = $true }
    }

    $ts = $Timestamp
    if ([string]::IsNullOrWhiteSpace($ts)) {
        $ts = (Get-Date).ToString('o')
    }

    $jobId = _QCT-FirstNonEmpty -Data $Data -Keys @('jobId', 'job_id')
    $docGuid = _QCT-FirstNonEmpty -Data $Data -Keys @('documentGuid', 'document_guid', 'docGuid', 'pw_objguid')
    $pkgId = _QCT-FirstNonEmpty -Data $Data -Keys @('sheetPackageId', 'sheet_package_id', 'packageId')
    $auditId = _QCT-FirstNonEmpty -Data $Data -Keys @('auditEventId', 'audit_event_id', 'id')
    $folder = _QCT-FirstNonEmpty -Data $Data -Keys @('folderPath', 'folder_path', 'resolved_folder', 'source_folder')

    $payload = @{
        ts = $ts
        level = $Level
        code = $Code
        message = $Message
        process_name = $proc
        run_id = $run
        data = $Data
    }
    $dataJson = ($payload | ConvertTo-Json -Depth 25 -Compress)
    $dataJson = _QDB-TruncateTelemetryPayload -Text $dataJson -MaxLength 32000

    if ([string]::IsNullOrWhiteSpace($DedupeKey)) {
        $DedupeKey = New-QCAutomationEventDedupeKey -Timestamp $ts -ProcessName $proc -Code $Code -Message $Message -JobId $jobId -RunId $run
    }

    $auditIdParam = $null
    if ($auditId -and $auditId -match '^\d+$') { $auditIdParam = [long]$auditId }

    try {
        $sql = @"
INSERT INTO automation_events (
    ts, process_name, run_id, level, code, message,
    job_id, document_guid, sheet_package_id, audit_event_id, folder_path,
    data_json, dedupe_key
) VALUES (
    @ts, @processName, @runId, @level, @code, @message,
    @jobId, @documentGuid, @sheetPackageId, @auditEventId, @folderPath,
    @dataJson, @dedupeKey
)
"@
        $params = @{
            ts = $ts
            processName = $proc
            runId = if ($run) { $run } else { [DBNull]::Value }
            level = $Level
            code = $Code
            message = if ($Message) { $Message } else { [DBNull]::Value }
            jobId = if ($jobId) { $jobId } else { [DBNull]::Value }
            documentGuid = if ($docGuid) { $docGuid } else { [DBNull]::Value }
            sheetPackageId = if ($pkgId) { $pkgId } else { [DBNull]::Value }
            auditEventId = if ($null -ne $auditIdParam) { $auditIdParam } else { [DBNull]::Value }
            folderPath = if ($folder) { $folder } else { [DBNull]::Value }
            dataJson = $dataJson
            dedupeKey = if ($DedupeKey) { $DedupeKey } else { [DBNull]::Value }
        }
        [void](Invoke-QCDatabaseNonQuery -Config $cfg -Sql $sql -Parameters $params)
        return New-QCSuccessResult -Code 'AUTOMATION_EVENT_WRITTEN' -Message 'Event persisted to automation_events.' -Data @{ written = $true; dedupeKey = $DedupeKey }
    } catch {
        $msg = [string]$_.Exception.Message
        if ($msg -match '(?i)duplicate|unique|2627|2601|UX_automation_events_dedupe') {
            return New-QCSuccessResult -Code 'AUTOMATION_EVENT_DUPLICATE' -Message 'Duplicate automation event skipped.' -Data @{ written = $false; duplicate = $true; dedupeKey = $DedupeKey }
        }
        return New-QCFailureResult -Code 'AUTOMATION_EVENT_WRITE_FAILED' -Message $msg -Data @{ written = $false; error = $msg }
    }
}

function Import-QCAutomationEventFromJsonLine {
    <#
    .SYNOPSIS
    Parses one JSONL log line and inserts into automation_events when eligible.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Line,
        [Parameter(Mandatory)][string]$ProcessName,
        [hashtable]$Config,
        [switch]$ForceProcessName
    )

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return New-QCSuccessResult -Code 'IMPORT_SKIPPED_EMPTY' -Message 'Empty line.' -Data @{ inserted = $false; skipped = $true }
    }
    try {
        $obj = $Line | ConvertFrom-Json
    } catch {
        return New-QCFailureResult -Code 'IMPORT_PARSE_FAILED' -Message $_.Exception.Message -Data @{ inserted = $false }
    }

    $data = @{}
    if ($obj.data) {
        $h = ConvertTo-HashtableDeep -Value $obj.data -ErrorAction SilentlyContinue
        if ($h -is [hashtable]) { $data = $h }
    }

    $proc = $ProcessName
    if (-not $ForceProcessName) {
        $fromData = _QCT-FirstNonEmpty -Data $data -Keys @('processName', 'process_name')
        if ($fromData) { $proc = $fromData }
    }

    $res = Write-QCAutomationEvent `
        -Level ([string]$obj.level) `
        -Code ([string]$obj.code) `
        -Message ([string]$obj.message) `
        -Data $data `
        -Timestamp ([string]$obj.ts) `
        -Config $Config `
        -ProcessName $proc `
        -RunId (_QCT-FirstNonEmpty -Data $data -Keys @('runId', 'run_id', 'pollRunId', 'poll_run_id'))

    $inserted = $res.IsSuccess -and $res.Data -and $res.Data.written
    $skipped = $res.IsSuccess -and $res.Data -and $res.Data.skipped
    $duplicate = $res.IsSuccess -and $res.Data -and $res.Data.detail -and $res.Data.detail.Code -eq 'AUTOMATION_EVENT_DUPLICATE'
    if ($duplicate) { $skipped = $true }
    $code = if ($inserted) { 'IMPORT_INSERTED' } elseif ($duplicate) { 'IMPORT_SKIPPED_DUPLICATE' } elseif ($skipped) { 'IMPORT_SKIPPED_FILTER' } else { 'IMPORT_NOT_INSERTED' }
    return New-QCSuccessResult -Code $code `
        -Message $res.Message -Data @{ inserted = [bool]$inserted; skipped = [bool]$skipped; duplicate = [bool]$duplicate; detail = $res }
}

Export-ModuleMember -Function @(
    'Set-QCAutomationTelemetryContext'
    'Get-QCAutomationTelemetrySettings'
    'Test-QCAutomationEventPersist'
    'New-QCAutomationEventDedupeKey'
    'Write-QCAutomationEvent'
    'Import-QCAutomationEventFromJsonLine'
)
