# Core.Database.psm1
# Responsibility: SQL Server connectivity and schema management for QC pipeline telemetry.
# The database is the reporting/control layer. The JSON queue remains the execution source.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

function _QDB-GetConnectionString {
    param([hashtable]$Config)
    if (-not $Config -or -not $Config.database -or -not $Config.database.connectionString) { return $null }
    return [string]$Config.database.connectionString
}

function _QDB-GetTimeout {
    param([hashtable]$Config, [int]$Override)
    if ($Override -gt 0) { return $Override }
    if ($Config -and $Config.database -and $Config.database.commandTimeout) { return [int]$Config.database.commandTimeout }
    return 60
}

function _QDB-IsEnabled {
    param([hashtable]$Config)
    if (-not $Config -or -not $Config.database) { return $false }
    try { return [bool]$Config.database.enabled } catch { return $false }
}

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

function Test-QCDatabaseEnabled {
    <#
    .SYNOPSIS
    Returns $true if database telemetry is enabled in config.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    return (_QDB-IsEnabled -Config $Config)
}

function Get-QCDatabaseConnection {
    <#
    .SYNOPSIS
    Opens and returns a SqlConnection. Caller must close/dispose.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $connStr = _QDB-GetConnectionString -Config $Config
    if (-not $connStr) {
        return New-QCFailureResult -Code 'DB_NO_CONNECTION_STRING' -Message 'database.connectionString not configured in appsettings.' -Data @{}
    }
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $conn.Open()
        return New-QCSuccessResult -Code 'DB_CONNECTED' -Message 'Database connection opened.' -Data @{ connection = $conn }
    } catch {
        return New-QCFailureResult -Code 'DB_CONNECTION_FAILED' -Message "Failed to connect: $($_.Exception.Message)" -Data @{ error = $_.Exception.Message }
    }
}

function Invoke-QCDatabaseQuery {
    <#
    .SYNOPSIS
    Executes a parameterized SELECT query and returns a DataTable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Parameters = @{},
        [int]$CommandTimeout = -1
    )

    $connStr = _QDB-GetConnectionString -Config $Config
    if (-not $connStr) {
        return New-QCFailureResult -Code 'DB_NO_CONNECTION_STRING' -Message 'database.connectionString not configured.' -Data @{}
    }
    $timeout = _QDB-GetTimeout -Config $Config -Override $CommandTimeout
    $conn = $null
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $timeout
        foreach ($key in $Parameters.Keys) {
            $val = $Parameters[$key]
            if ($null -eq $val) { $val = [DBNull]::Value }
            [void]$cmd.Parameters.AddWithValue("@$key", $val)
        }
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $table = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
        return New-QCSuccessResult -Code 'DB_QUERY_OK' -Message "Query returned $($table.Rows.Count) rows." -Data @{ table = $table; rowCount = $table.Rows.Count }
    } catch {
        return New-QCFailureResult -Code 'DB_QUERY_FAILED' -Message "Query failed: $($_.Exception.Message)" -Data @{ sql = $Sql; error = $_.Exception.Message }
    } finally {
        if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
        if ($conn) { $conn.Dispose() }
    }
}

function Invoke-QCDatabaseNonQuery {
    <#
    .SYNOPSIS
    Executes a parameterized INSERT/UPDATE/DELETE and returns rows affected.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Parameters = @{},
        [int]$CommandTimeout = -1
    )

    $connStr = _QDB-GetConnectionString -Config $Config
    if (-not $connStr) {
        return New-QCFailureResult -Code 'DB_NO_CONNECTION_STRING' -Message 'database.connectionString not configured.' -Data @{}
    }
    $timeout = _QDB-GetTimeout -Config $Config -Override $CommandTimeout
    $conn = $null
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $timeout
        foreach ($key in $Parameters.Keys) {
            $val = $Parameters[$key]
            if ($null -eq $val) { $val = [DBNull]::Value }
            [void]$cmd.Parameters.AddWithValue("@$key", $val)
        }
        $rowsAffected = $cmd.ExecuteNonQuery()
        return New-QCSuccessResult -Code 'DB_NONQUERY_OK' -Message "$rowsAffected rows affected." -Data @{ rowsAffected = $rowsAffected }
    } catch {
        return New-QCFailureResult -Code 'DB_NONQUERY_FAILED' -Message "Non-query failed: $($_.Exception.Message)" -Data @{ sql = $Sql; error = $_.Exception.Message }
    } finally {
        if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
        if ($conn) { $conn.Dispose() }
    }
}

function Invoke-QCDatabaseScalar {
    <#
    .SYNOPSIS
    Executes a parameterized scalar query and returns a single value.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Parameters = @{},
        [int]$CommandTimeout = -1
    )

    $connStr = _QDB-GetConnectionString -Config $Config
    if (-not $connStr) {
        return New-QCFailureResult -Code 'DB_NO_CONNECTION_STRING' -Message 'database.connectionString not configured.' -Data @{}
    }
    $timeout = _QDB-GetTimeout -Config $Config -Override $CommandTimeout
    $conn = $null
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $timeout
        foreach ($key in $Parameters.Keys) {
            $val = $Parameters[$key]
            if ($null -eq $val) { $val = [DBNull]::Value }
            [void]$cmd.Parameters.AddWithValue("@$key", $val)
        }
        $result = $cmd.ExecuteScalar()
        if ($result -is [DBNull]) { $result = $null }
        return New-QCSuccessResult -Code 'DB_SCALAR_OK' -Message 'Scalar query executed.' -Data @{ value = $result }
    } catch {
        return New-QCFailureResult -Code 'DB_SCALAR_FAILED' -Message "Scalar query failed: $($_.Exception.Message)" -Data @{ sql = $Sql; error = $_.Exception.Message }
    } finally {
        if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
        if ($conn) { $conn.Dispose() }
    }
}

function Invoke-QCDatabaseBatch {
    <#
    .SYNOPSIS
    Executes multiple SQL statements separated by GO (like sqlcmd).
    Used for schema initialization where CREATE VIEW must be in its own batch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Sql,
        [int]$CommandTimeout = -1
    )

    $connStr = _QDB-GetConnectionString -Config $Config
    if (-not $connStr) {
        return New-QCFailureResult -Code 'DB_NO_CONNECTION_STRING' -Message 'database.connectionString not configured.' -Data @{}
    }
    $timeout = _QDB-GetTimeout -Config $Config -Override $CommandTimeout
    $conn = $null
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        $conn.Open()

        $batches = [regex]::Split($Sql, '^\s*GO\s*$', [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $executed = 0
        foreach ($batch in $batches) {
            $trimmed = $batch.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $trimmed
            $cmd.CommandTimeout = $timeout
            [void]$cmd.ExecuteNonQuery()
            $executed++
        }
        return New-QCSuccessResult -Code 'DB_BATCH_OK' -Message "$executed batches executed." -Data @{ batchCount = $executed }
    } catch {
        return New-QCFailureResult -Code 'DB_BATCH_FAILED' -Message "Batch execution failed: $($_.Exception.Message)" -Data @{ error = $_.Exception.Message }
    } finally {
        if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
        if ($conn) { $conn.Dispose() }
    }
}

# ---------------------------------------------------------------------------
# Schema management
# ---------------------------------------------------------------------------

function Initialize-QCDatabaseSchema {
    <#
    .SYNOPSIS
    Creates all QC pipeline telemetry tables, indexes, and views idempotently.
    Tracks applied versions in schema_version table.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'DB_DISABLED' -Message 'Database is not enabled in config.' -Data @{}
    }

    $targetVersion = '1.0.0'
    $schemaSql = _QDB-GetSchemaV1

    $connRes = Get-QCDatabaseConnection -Config $Config
    if (-not $connRes.IsSuccess) { return $connRes }
    $conn = $connRes.Data.connection

    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "IF OBJECT_ID('dbo.schema_version', 'U') IS NOT NULL SELECT MAX(version) FROM schema_version ELSE SELECT NULL"
        $currentVersion = $cmd.ExecuteScalar()
        if ($currentVersion -is [DBNull]) { $currentVersion = $null }

        if ($currentVersion -eq $targetVersion) {
            return New-QCSuccessResult -Code 'DB_SCHEMA_CURRENT' -Message "Schema already at version $targetVersion." -Data @{ version = $targetVersion }
        }

        $batches = [regex]::Split($schemaSql, '^\s*GO\s*$', [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $executed = 0
        foreach ($batch in $batches) {
            $trimmed = $batch.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            $batchCmd = $conn.CreateCommand()
            $batchCmd.CommandText = $trimmed
            $batchCmd.CommandTimeout = 120
            [void]$batchCmd.ExecuteNonQuery()
            $executed++
        }

        $insertCmd = $conn.CreateCommand()
        $insertCmd.CommandText = "INSERT INTO schema_version (version, description) VALUES (@version, @desc)"
        [void]$insertCmd.Parameters.AddWithValue("@version", $targetVersion)
        [void]$insertCmd.Parameters.AddWithValue("@desc", "Initial telemetry schema: audit_events, document_activity, document_state_history, transition_events, poll_runs, processing_jobs, notification_log")
        [void]$insertCmd.ExecuteNonQuery()

        return New-QCSuccessResult -Code 'DB_SCHEMA_INITIALIZED' -Message "Schema initialized to version $targetVersion ($executed batches)." -Data @{ version = $targetVersion; batchCount = $executed }
    } catch {
        return New-QCFailureResult -Code 'DB_SCHEMA_FAILED' -Message "Schema initialization failed: $($_.Exception.Message)" -Data @{ error = $_.Exception.Message }
    } finally {
        if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
        if ($conn) { $conn.Dispose() }
    }
}

function _QDB-GetSchemaV1 {
    return @'

-- schema_version: tracks applied migrations
IF OBJECT_ID('dbo.schema_version', 'U') IS NULL
CREATE TABLE schema_version (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    version         NVARCHAR(50) NOT NULL,
    applied_at      DATETIME2(3) NOT NULL DEFAULT SYSDATETIME(),
    description     NVARCHAR(500)
);

GO

-- audit_events: raw PW audit trail records from dms_audt
IF OBJECT_ID('dbo.audit_events', 'U') IS NULL
CREATE TABLE audit_events (
    id              BIGINT IDENTITY(1,1) PRIMARY KEY,
    captured_at     DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    poll_run_id     INT,
    pw_acttime      NVARCHAR(50) NOT NULL,
    pw_action       INT NOT NULL,
    pw_action_name  NVARCHAR(100),
    pw_objtype      INT NOT NULL DEFAULT 0,
    pw_objno        INT,
    pw_objguid      NVARCHAR(50),
    pw_parentguid   NVARCHAR(50),
    pw_userno       INT,
    pw_itemname     NVARCHAR(500),
    pw_itemdesc     NVARCHAR(1000),
    pw_textparam    NVARCHAR(2000),
    resolved_folder NVARCHAR(1000),
    candidate_type  NVARCHAR(50),
    processed       BIT NOT NULL DEFAULT 0,
    enqueued_job_id NVARCHAR(200)
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_audit_events_acttime')
    CREATE INDEX idx_audit_events_acttime ON audit_events(pw_acttime);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_audit_events_objguid')
    CREATE INDEX idx_audit_events_objguid ON audit_events(pw_objguid);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_audit_events_folder')
    CREATE INDEX idx_audit_events_folder ON audit_events(resolved_folder);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_audit_events_processed')
    CREATE INDEX idx_audit_events_processed ON audit_events(processed);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_audit_events_poll_run')
    CREATE INDEX idx_audit_events_poll_run ON audit_events(poll_run_id);

GO

-- document_activity: enriched per-document summary (upserted)
IF OBJECT_ID('dbo.document_activity', 'U') IS NULL
CREATE TABLE document_activity (
    id                  INT IDENTITY(1,1) PRIMARY KEY,
    document_guid       NVARCHAR(50) NOT NULL,
    document_name       NVARCHAR(500),
    folder_path         NVARCHAR(1000),
    last_action         NVARCHAR(100),
    last_action_code    INT,
    last_action_time    NVARCHAR(50),
    last_action_user    INT,
    total_events        INT NOT NULL DEFAULT 0,
    first_seen          DATETIMEOFFSET(3),
    last_seen           DATETIMEOFFSET(3),
    CONSTRAINT uq_doc_activity_guid UNIQUE(document_guid)
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_doc_activity_folder')
    CREATE INDEX idx_doc_activity_folder ON document_activity(folder_path);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_doc_activity_lastaction')
    CREATE INDEX idx_doc_activity_lastaction ON document_activity(last_action_time);

GO

-- document_state_history: time-series of state/attribute changes
IF OBJECT_ID('dbo.document_state_history', 'U') IS NULL
CREATE TABLE document_state_history (
    id              BIGINT IDENTITY(1,1) PRIMARY KEY,
    document_guid   NVARCHAR(50) NOT NULL,
    document_name   NVARCHAR(500),
    folder_path     NVARCHAR(1000),
    captured_at     DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    event_type      NVARCHAR(50),
    source_audit_id BIGINT,
    old_value       NVARCHAR(500),
    new_value       NVARCHAR(500),
    field_name      NVARCHAR(200),
    changed_by_user INT
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_state_hist_docguid')
    CREATE INDEX idx_state_hist_docguid ON document_state_history(document_guid);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_state_hist_captured')
    CREATE INDEX idx_state_hist_captured ON document_state_history(captured_at);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_state_hist_type')
    CREATE INDEX idx_state_hist_type ON document_state_history(event_type);

GO

-- transition_events: business-level events (QC stage changes, check-ins)
IF OBJECT_ID('dbo.transition_events', 'U') IS NULL
CREATE TABLE transition_events (
    id                  INT IDENTITY(1,1) PRIMARY KEY,
    detected_at         DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    document_guid       NVARCHAR(50) NOT NULL,
    document_name       NVARCHAR(500),
    folder_path         NVARCHAR(1000),
    transition_type     NVARCHAR(100) NOT NULL,
    from_value          NVARCHAR(500),
    to_value            NVARCHAR(500),
    trigger_audit_id    BIGINT,
    job_id              NVARCHAR(200),
    job_type            NVARCHAR(50),
    notification_sent   BIT NOT NULL DEFAULT 0,
    notification_id     NVARCHAR(200)
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_transition_docguid')
    CREATE INDEX idx_transition_docguid ON transition_events(document_guid);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_transition_type')
    CREATE INDEX idx_transition_type ON transition_events(transition_type);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_transition_detected')
    CREATE INDEX idx_transition_detected ON transition_events(detected_at);

GO

-- poll_runs: operational health of the audit poller
IF OBJECT_ID('dbo.poll_runs', 'U') IS NULL
CREATE TABLE poll_runs (
    id                  INT IDENTITY(1,1) PRIMARY KEY,
    started_at          DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    completed_at        DATETIMEOFFSET(3),
    watermark_before    NVARCHAR(50),
    watermark_after     NVARCHAR(50),
    events_fetched      INT NOT NULL DEFAULT 0,
    events_relevant     INT NOT NULL DEFAULT 0,
    candidates_created  INT NOT NULL DEFAULT 0,
    jobs_enqueued       INT NOT NULL DEFAULT 0,
    duration_ms         INT,
    error_message       NVARCHAR(2000),
    is_reconciliation   BIT NOT NULL DEFAULT 0
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_poll_runs_started')
    CREATE INDEX idx_poll_runs_started ON poll_runs(started_at);

GO

-- processing_jobs: mirrors queue job outcomes for dashboards (read-only telemetry)
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NULL
CREATE TABLE processing_jobs (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    job_id          NVARCHAR(200) NOT NULL,
    job_type        NVARCHAR(50) NOT NULL,
    source_path     NVARCHAR(1000),
    source_folder   NVARCHAR(1000),
    created_at      DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    started_at      DATETIMEOFFSET(3),
    completed_at    DATETIMEOFFSET(3),
    status          NVARCHAR(20) NOT NULL DEFAULT 'pending',
    trigger_source  NVARCHAR(50),
    trigger_audit_id BIGINT,
    dedupe_key      NVARCHAR(200),
    attempt_count   INT NOT NULL DEFAULT 0,
    duration_ms     INT,
    error_code      NVARCHAR(100),
    error_message   NVARCHAR(2000),
    result_data     NVARCHAR(MAX),
    CONSTRAINT uq_processing_jobs_jobid UNIQUE(job_id)
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_status')
    CREATE INDEX idx_jobs_status ON processing_jobs(status);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_type')
    CREATE INDEX idx_jobs_type ON processing_jobs(job_type);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_created')
    CREATE INDEX idx_jobs_created ON processing_jobs(created_at);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_folder')
    CREATE INDEX idx_jobs_folder ON processing_jobs(source_folder);

GO

-- notification_log: sent notification tracking and dedupe
IF OBJECT_ID('dbo.notification_log', 'U') IS NULL
CREATE TABLE notification_log (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    sent_at         DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    event_type      NVARCHAR(100) NOT NULL,
    document_guid   NVARCHAR(50),
    document_name   NVARCHAR(500),
    folder_path     NVARCHAR(1000),
    recipients      NVARCHAR(2000),
    subject         NVARCHAR(500),
    dedupe_key      NVARCHAR(500),
    provider        NVARCHAR(50),
    success         BIT NOT NULL DEFAULT 1,
    error_message   NVARCHAR(2000),
    transition_id   INT
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_notif_dedupekey')
    CREATE INDEX idx_notif_dedupekey ON notification_log(dedupe_key);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_notif_sent')
    CREATE INDEX idx_notif_sent ON notification_log(sent_at);

GO

-- v_qc_cycle_aging: active QC cycle duration per document
IF OBJECT_ID('dbo.v_qc_cycle_aging', 'V') IS NOT NULL DROP VIEW v_qc_cycle_aging;

GO

CREATE VIEW v_qc_cycle_aging AS
SELECT
    document_guid,
    document_name,
    folder_path,
    MIN(CASE WHEN event_type = 'STATE_CHANGE' AND new_value LIKE '%QC Received%' THEN captured_at END) AS qc_received_at,
    MAX(captured_at) AS last_activity_at,
    ROUND(DATEDIFF(MINUTE,
        MIN(CASE WHEN event_type = 'STATE_CHANGE' AND new_value LIKE '%QC Received%' THEN captured_at END),
        GETDATE()) / 60.0, 1) AS hours_in_qc
FROM document_state_history
GROUP BY document_guid, document_name, folder_path;

GO

-- v_folder_activity: folder event counts and last activity (7 days)
IF OBJECT_ID('dbo.v_folder_activity', 'V') IS NOT NULL DROP VIEW v_folder_activity;

GO

CREATE VIEW v_folder_activity AS
SELECT
    resolved_folder AS folder_path,
    COUNT(*) AS event_count,
    COUNT(DISTINCT pw_objguid) AS unique_documents,
    MAX(pw_acttime) AS last_activity,
    MIN(pw_acttime) AS first_activity
FROM audit_events
WHERE captured_at > DATEADD(DAY, -7, SYSDATETIMEOFFSET())
GROUP BY resolved_folder;

GO

-- v_poller_health: recent poll run status
IF OBJECT_ID('dbo.v_poller_health', 'V') IS NOT NULL DROP VIEW v_poller_health;

GO

CREATE VIEW v_poller_health AS
SELECT TOP 100
    id,
    started_at,
    duration_ms,
    events_fetched,
    events_relevant,
    jobs_enqueued,
    CASE WHEN error_message IS NOT NULL THEN 'ERROR' ELSE 'OK' END AS run_status,
    watermark_after
FROM poll_runs
ORDER BY started_at DESC;

GO

-- v_job_summary: processing job counts by type/status
IF OBJECT_ID('dbo.v_job_summary', 'V') IS NOT NULL DROP VIEW v_job_summary;

GO

CREATE VIEW v_job_summary AS
SELECT
    job_type,
    status,
    COUNT(*) AS job_count,
    AVG(duration_ms) AS avg_duration_ms,
    MAX(completed_at) AS last_completed
FROM processing_jobs
GROUP BY job_type, status;

'@
}

Export-ModuleMember -Function Test-QCDatabaseEnabled, Get-QCDatabaseConnection, Invoke-QCDatabaseQuery, Invoke-QCDatabaseNonQuery, Invoke-QCDatabaseScalar, Invoke-QCDatabaseBatch, Initialize-QCDatabaseSchema
