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

    $targetVersion = '1.2.0'
    $schemaV1 = _QDB-GetSchemaV1
    $schemaV1_1 = _QDB-GetSchemaV1dot1
    $schemaV1_2 = _QDB-GetSchemaV1dot2
    $schemaSql = $schemaV1 + [Environment]::NewLine + $schemaV1_1 + [Environment]::NewLine + $schemaV1_2

    $connRes = Get-QCDatabaseConnection -Config $Config
    if (-not $connRes.IsSuccess) { return $connRes }
    $conn = $connRes.Data.connection

    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "IF OBJECT_ID('dbo.schema_version', 'U') IS NOT NULL SELECT MAX(version) FROM schema_version ELSE SELECT NULL"
        $currentVersion = $cmd.ExecuteScalar()
        if ($currentVersion -is [DBNull]) { $currentVersion = $null }

        if ($currentVersion -eq $targetVersion) {
            $patchSql = _QDB-GetSchemaV1dot2Additive
            $patchBatches = [regex]::Split($patchSql, '^\s*GO\s*$', [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $patchExecuted = 0
            foreach ($batch in $patchBatches) {
                $trimmed = $batch.Trim()
                if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
                $batchCmd = $conn.CreateCommand()
                $batchCmd.CommandText = $trimmed
                $batchCmd.CommandTimeout = 120
                [void]$batchCmd.ExecuteNonQuery()
                $patchExecuted++
            }
            return New-QCSuccessResult -Code 'DB_SCHEMA_CURRENT' -Message "Schema already at version $targetVersion (additive patches: $patchExecuted)." -Data @{ version = $targetVersion; patchCount = $patchExecuted }
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
        [void]$insertCmd.Parameters.AddWithValue("@desc", "QC telemetry schema through comment sync tables (qc_comment_*, qc_workflow_events)")
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

function _QDB-GetSchemaV1dot1 {
    return @'

GO

-- sheet_index: tracks all sheets in watched Sheets folders with QC PDF pairing and ownership
IF OBJECT_ID('dbo.sheet_index', 'U') IS NULL
CREATE TABLE sheet_index (
    id                  INT IDENTITY(1,1) PRIMARY KEY,
    document_guid       NVARCHAR(40) NOT NULL,
    document_name       NVARCHAR(500) NOT NULL,
    document_number     INT NULL,
    folder_path         NVARCHAR(1000) NOT NULL,
    project_name        NVARCHAR(200) NULL,
    watch_root          NVARCHAR(500) NULL,
    extension           NVARCHAR(20) NULL,

    qc_pdf_guid         NVARCHAR(40) NULL,
    qc_pdf_name         NVARCHAR(500) NULL,
    source_type         NVARCHAR(10) NULL,

    designer_email      NVARCHAR(200) NULL,
    reviewer_email      NVARCHAR(200) NULL,

    pw_state_name       NVARCHAR(100) NULL,
    qc_stage            NVARCHAR(20) NULL,
    qc_status           NVARCHAR(50) NULL,

    first_seen_at       DATETIMEOFFSET NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    last_updated_at     DATETIMEOFFSET NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    last_audit_event_at DATETIMEOFFSET NULL,
    file_modified_at    DATETIMEOFFSET NULL,

    CONSTRAINT UQ_sheet_index_doc_guid UNIQUE (document_guid)
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_index_folder')
    CREATE INDEX IX_sheet_index_folder ON sheet_index (folder_path);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_index_qc_pair')
    CREATE INDEX IX_sheet_index_qc_pair ON sheet_index (qc_pdf_guid) WHERE qc_pdf_guid IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_index_designer')
    CREATE INDEX IX_sheet_index_designer ON sheet_index (designer_email) WHERE designer_email IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_index_reviewer')
    CREATE INDEX IX_sheet_index_reviewer ON sheet_index (reviewer_email) WHERE reviewer_email IS NOT NULL;

GO

-- v_sheet_status: project status overview of all indexed sheets
IF OBJECT_ID('dbo.v_sheet_status', 'V') IS NOT NULL DROP VIEW v_sheet_status;

GO

CREATE VIEW v_sheet_status AS
SELECT
    s.document_name,
    s.folder_path,
    s.project_name,
    s.extension,
    s.designer_email,
    s.reviewer_email,
    s.pw_state_name,
    s.qc_stage,
    s.qc_status,
    s.qc_pdf_name,
    CASE WHEN s.qc_pdf_guid IS NOT NULL THEN 1 ELSE 0 END AS has_qc_pdf,
    s.last_updated_at,
    s.file_modified_at,
    s.watch_root
FROM sheet_index s;

'@
}

function _QDB-GetSchemaV1dot2 {
    return @'

GO

-- qc_comment_runs: one processor execution per *-qc.pdf sync job
IF OBJECT_ID('dbo.qc_comment_runs', 'U') IS NULL
CREATE TABLE qc_comment_runs (
    run_id                  BIGINT IDENTITY(1,1) PRIMARY KEY,
    job_id                  NVARCHAR(100) NOT NULL,
    document_id             NVARCHAR(40) NULL,
    project_id              NVARCHAR(200) NULL,
    pw_path                 NVARCHAR(1000) NULL,
    file_name               NVARCHAR(500) NULL,
    file_hash               NVARCHAR(128) NULL,
    source_modified_utc     DATETIMEOFFSET(3) NULL,
    previous_pw_state       NVARCHAR(100) NULL,
    target_pw_state         NVARCHAR(100) NULL,
    state_update_result     NVARCHAR(500) NULL,
    parser_status           NVARCHAR(50) NULL,
    processor_version       NVARCHAR(50) NULL,
    created_utc             DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_runs_job')
    CREATE INDEX IX_qc_comment_runs_job ON qc_comment_runs (job_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_runs_doc')
    CREATE INDEX IX_qc_comment_runs_doc ON qc_comment_runs (document_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_runs_created')
    CREATE INDEX IX_qc_comment_runs_created ON qc_comment_runs (created_utc);

GO

-- qc_comments: annotation snapshot per run
IF OBJECT_ID('dbo.qc_comments', 'U') IS NULL
CREATE TABLE qc_comments (
    comment_record_id       BIGINT IDENTITY(1,1) PRIMARY KEY,
    run_id                  BIGINT NOT NULL,
    document_id             NVARCHAR(40) NULL,
    annotation_id           NVARCHAR(50) NOT NULL,
    page_number             INT NULL,
    author                  NVARCHAR(200) NULL,
    subject                 NVARCHAR(500) NULL,
    comment_text            NVARCHAR(MAX) NULL,
    color                   NVARCHAR(100) NULL,
    status                  NVARCHAR(100) NULL,
    status_author           NVARCHAR(200) NULL,
    status_timestamp_utc    DATETIMEOFFSET(3) NULL,
    created_utc             DATETIMEOFFSET(3) NULL,
    modified_utc            DATETIMEOFFSET(3) NULL,
    parent_annotation_id    NVARCHAR(50) NULL,
    raw_json                NVARCHAR(MAX) NULL,
    inserted_utc            DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comments_run')
    CREATE INDEX IX_qc_comments_run ON qc_comments (run_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comments_doc_annot')
    CREATE INDEX IX_qc_comments_doc_annot ON qc_comments (document_id, annotation_id);

GO

-- qc_comment_status_history: status transitions over time
IF OBJECT_ID('dbo.qc_comment_status_history', 'U') IS NULL
CREATE TABLE qc_comment_status_history (
    history_id                  BIGINT IDENTITY(1,1) PRIMARY KEY,
    document_id                 NVARCHAR(40) NULL,
    annotation_id               NVARCHAR(50) NOT NULL,
    previous_status             NVARCHAR(100) NULL,
    current_status              NVARCHAR(100) NULL,
    previous_status_timestamp_utc DATETIMEOFFSET(3) NULL,
    current_status_timestamp_utc  DATETIMEOFFSET(3) NULL,
    detected_run_id             BIGINT NULL,
    detected_utc                DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_status_hist_doc')
    CREATE INDEX IX_qc_comment_status_hist_doc ON qc_comment_status_history (document_id, annotation_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_status_hist_run')
    CREATE INDEX IX_qc_comment_status_hist_run ON qc_comment_status_history (detected_run_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_status_hist_detected')
    CREATE INDEX IX_qc_comment_status_hist_detected ON qc_comment_status_history (detected_utc);

GO

-- qc_workflow_events: workflow transition audit log (separate from comment snapshots)
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NULL
CREATE TABLE qc_workflow_events (
    event_id                BIGINT IDENTITY(1,1) PRIMARY KEY,
    run_id                  BIGINT NULL,
    job_id                  NVARCHAR(100) NULL,
    document_id             NVARCHAR(40) NULL,
    event_type              NVARCHAR(100) NOT NULL,
    previous_pw_state       NVARCHAR(100) NULL,
    target_pw_state         NVARCHAR(100) NULL,
    decision_code           NVARCHAR(100) NULL,
    processor_version       NVARCHAR(50) NULL,
    payload_json            NVARCHAR(MAX) NULL,
    created_utc             DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_run')
    CREATE INDEX IX_qc_workflow_events_run ON qc_workflow_events (run_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_doc')
    CREATE INDEX IX_qc_workflow_events_doc ON qc_workflow_events (document_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_job')
    CREATE INDEX IX_qc_workflow_events_job ON qc_workflow_events (job_id);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_created')
    CREATE INDEX IX_qc_workflow_events_created ON qc_workflow_events (created_utc);

'@
}

function _QDB-GetSchemaV1dot2Additive {
    return @'
GO
IF OBJECT_ID('dbo.qc_comment_runs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_runs_created')
    CREATE INDEX IX_qc_comment_runs_created ON qc_comment_runs (created_utc);
IF OBJECT_ID('dbo.qc_comment_status_history', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_status_hist_run')
    CREATE INDEX IX_qc_comment_status_hist_run ON qc_comment_status_history (detected_run_id);
IF OBJECT_ID('dbo.qc_comment_status_history', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_comment_status_hist_detected')
    CREATE INDEX IX_qc_comment_status_hist_detected ON qc_comment_status_history (detected_utc);
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_job')
    CREATE INDEX IX_qc_workflow_events_job ON qc_workflow_events (job_id);
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_created')
    CREATE INDEX IX_qc_workflow_events_created ON qc_workflow_events (created_utc);
'@
}

function Test-QCDatabaseWritesAllowed {
    <#
    .SYNOPSIS
    Returns $true when config allows mutating DB writes (respects dry-run).
    #>
    [CmdletBinding()]
    param([hashtable]$Config)
    if (-not (_QDB-IsEnabled -Config $Config)) { return $false }
    $dryRun = $false
    if ($Config.ContainsKey('dryRun')) { try { $dryRun = [bool]$Config.dryRun } catch { } }
    if ($dryRun) {
        $allow = $false
        if ($Config.database -and $null -ne $Config.database.allowWritesInDryRun) {
            try { $allow = [bool]$Config.database.allowWritesInDryRun } catch { $allow = $false }
        }
        return $allow
    }
    return $true
}

# ---------------------------------------------------------------------------
# Fire-and-forget telemetry writers
# These silently no-op when the database is disabled or unreachable.
# Pipeline execution must NEVER fail because telemetry fails.
# ---------------------------------------------------------------------------

function _QDB-SafeWrite {
    param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters)
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    try { Invoke-QCDatabaseNonQuery -Config $Config -Sql $Sql -Parameters $Parameters | Out-Null }
    catch { }
}

function Get-QCProcessingJobType {
    <#
    .SYNOPSIS
    Maps queue/processor job types to processing_jobs.job_type for dashboards and reporting.
  .DESCRIPTION
    Queue jobs may use granular types (e.g. QC_COMMENT_STATUS_SYNC) while processing_jobs
    uses consolidated reporting types (e.g. QC_STATE).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$QueueJobType,
        [hashtable]$Config = $null
    )

    $map = @{
        'QC_COMMENT_STATUS_SYNC' = 'QC_STATE'
    }
    if ($Config -and $Config.ContainsKey('database') -and $Config.database -and $Config.database.processingJobTypeMap) {
        $custom = $Config.database.processingJobTypeMap
        if ($custom -is [hashtable]) {
            foreach ($k in $custom.Keys) { $map[[string]$k] = [string]$custom[$k] }
        }
    }

    $key = [string]$QueueJobType
    if ($map.ContainsKey($key)) { return [string]$map[$key] }
    return $key
}

function Write-QCJobTelemetry {
    <#
    .SYNOPSIS
    Upserts a processing job outcome into the processing_jobs table.
    Fire-and-forget: silently no-ops if DB is disabled or unreachable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][string]$JobType,
        [Parameter(Mandatory)][string]$Status,
        [string]$SourcePath,
        [string]$SourceFolder,
        [string]$DedupeKey,
        [string]$TriggerSource,
        [int]$AttemptCount = 0,
        [Nullable[int]]$DurationMs,
        [string]$ErrorCode,
        [string]$ErrorMessage,
        [string]$ResultData
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    $telemetryJobType = Get-QCProcessingJobType -QueueJobType $JobType -Config $Config
    try {
        $sql = @"
MERGE processing_jobs AS tgt
USING (SELECT @jobId AS job_id) AS src ON tgt.job_id = src.job_id
WHEN MATCHED THEN UPDATE SET
    status = @status,
    completed_at = CASE WHEN @status IN ('succeeded','failed') THEN SYSDATETIMEOFFSET() ELSE tgt.completed_at END,
    attempt_count = @attemptCount,
    duration_ms = @durationMs,
    error_code = @errorCode,
    error_message = @errorMessage,
    result_data = @resultData
WHEN NOT MATCHED THEN INSERT
    (job_id, job_type, status, source_path, source_folder, dedupe_key, trigger_source, attempt_count, duration_ms, error_code, error_message, result_data)
VALUES
    (@jobId, @jobType, @status, @sourcePath, @sourceFolder, @dedupeKey, @triggerSource, @attemptCount, @durationMs, @errorCode, @errorMessage, @resultData);
"@
        $params = @{
            jobId         = $JobId
            jobType       = $telemetryJobType
            status        = $Status
            sourcePath    = if ($SourcePath)    { $SourcePath }    else { $null }
            sourceFolder  = if ($SourceFolder)  { $SourceFolder }  else { $null }
            dedupeKey     = if ($DedupeKey)      { $DedupeKey }      else { $null }
            triggerSource = if ($TriggerSource) { $TriggerSource } else { $null }
            attemptCount  = $AttemptCount
            durationMs    = if ($null -ne $DurationMs) { $DurationMs } else { $null }
            errorCode     = if ($ErrorCode)     { $ErrorCode }     else { $null }
            errorMessage  = if ($ErrorMessage)  { $ErrorMessage }  else { $null }
            resultData    = if ($ResultData)    { $ResultData }    else { $null }
        }
        Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
    } catch { }
}

function Write-QCPollRunTelemetry {
    <#
    .SYNOPSIS
    Inserts a completed scan/poll run into the poll_runs table.
    Fire-and-forget.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [int]$EventsFetched = 0,
        [int]$EventsRelevant = 0,
        [int]$CandidatesCreated = 0,
        [int]$JobsEnqueued = 0,
        [Nullable[int]]$DurationMs,
        [string]$ErrorMessage,
        [string]$WatermarkBefore,
        [string]$WatermarkAfter,
        [bool]$IsReconciliation = $false
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    try {
        $sql = @"
INSERT INTO poll_runs
    (started_at, completed_at, watermark_before, watermark_after, events_fetched, events_relevant, candidates_created, jobs_enqueued, duration_ms, error_message, is_reconciliation)
VALUES
    (DATEADD(MILLISECOND, -@durationMs, SYSDATETIMEOFFSET()), SYSDATETIMEOFFSET(), @watermarkBefore, @watermarkAfter, @eventsFetched, @eventsRelevant, @candidatesCreated, @jobsEnqueued, @durationMs, @errorMessage, @isReconciliation)
"@
        $params = @{
            eventsFetched     = $EventsFetched
            eventsRelevant    = $EventsRelevant
            candidatesCreated = $CandidatesCreated
            jobsEnqueued      = $JobsEnqueued
            durationMs        = if ($null -ne $DurationMs) { $DurationMs } else { 0 }
            errorMessage      = if ($ErrorMessage) { $ErrorMessage } else { $null }
            watermarkBefore   = if ($WatermarkBefore) { $WatermarkBefore } else { $null }
            watermarkAfter    = if ($WatermarkAfter) { $WatermarkAfter } else { $null }
            isReconciliation  = if ($IsReconciliation) { 1 } else { 0 }
        }
        Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
    } catch { }
}

function Write-QCNotificationTelemetry {
    <#
    .SYNOPSIS
    Inserts a sent notification record into the notification_log table.
    Fire-and-forget.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$EventType,
        [string]$DocumentGuid,
        [string]$DocumentName,
        [string]$FolderPath,
        [string]$Recipients,
        [string]$Subject,
        [string]$DedupeKey,
        [string]$Provider,
        [bool]$Success = $true,
        [string]$ErrorMessage,
        [Nullable[int]]$TransitionId
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    try {
        $sql = @"
INSERT INTO notification_log
    (event_type, document_guid, document_name, folder_path, recipients, subject, dedupe_key, provider, success, error_message, transition_id)
VALUES
    (@eventType, @documentGuid, @documentName, @folderPath, @recipients, @subject, @dedupeKey, @provider, @success, @errorMessage, @transitionId)
"@
        $params = @{
            eventType    = $EventType
            documentGuid = if ($DocumentGuid) { $DocumentGuid } else { $null }
            documentName = if ($DocumentName) { $DocumentName } else { $null }
            folderPath   = if ($FolderPath)   { $FolderPath }   else { $null }
            recipients   = if ($Recipients)   { $Recipients }   else { $null }
            subject      = if ($Subject)      { $Subject }      else { $null }
            dedupeKey    = if ($DedupeKey)     { $DedupeKey }     else { $null }
            provider     = if ($Provider)      { $Provider }      else { $null }
            success      = if ($Success) { 1 } else { 0 }
            errorMessage = if ($ErrorMessage) { $ErrorMessage } else { $null }
            transitionId = if ($null -ne $TransitionId) { $TransitionId } else { $null }
        }
        Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
    } catch { }
}

function Write-QCSheetIndex {
    <#
    .SYNOPSIS
    Upserts a sheet into the sheet_index table. Fire-and-forget.
    Accepts either raw field values or a PW document object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [int]$DocumentNumber = 0,
        [string]$ProjectName,
        [string]$WatchRoot,
        [string]$Extension,
        [string]$SourceType,
        [string]$DesignerEmail,
        [string]$ReviewerEmail,
        [string]$PwStateName,
        [string]$QcStage,
        [string]$QcStatus,
        [string]$LastAuditEventAt,
        [string]$FileModifiedAt,
        [switch]$SetOwnershipFromProjectWise
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    try {
        if (-not $Extension -and $DocumentName) {
            $ext = [System.IO.Path]::GetExtension($DocumentName)
            if ($ext) { $Extension = $ext.ToLowerInvariant() }
        }
        $setOwnership = if ($SetOwnershipFromProjectWise) { 1 } else { 0 }
        $sql = @"
MERGE sheet_index AS tgt
USING (SELECT @docGuid AS document_guid) AS src ON tgt.document_guid = src.document_guid
WHEN MATCHED THEN UPDATE SET
    document_name = @docName,
    folder_path = @folderPath,
    document_number = CASE WHEN @docNumber > 0 THEN @docNumber ELSE tgt.document_number END,
    project_name = COALESCE(@projectName, tgt.project_name),
    watch_root = COALESCE(@watchRoot, tgt.watch_root),
    extension = COALESCE(@extension, tgt.extension),
    source_type = COALESCE(@sourceType, tgt.source_type),
    designer_email = CASE WHEN @setOwnership = 1 THEN @designerEmail ELSE COALESCE(@designerEmail, tgt.designer_email) END,
    reviewer_email = CASE WHEN @setOwnership = 1 THEN @reviewerEmail ELSE COALESCE(@reviewerEmail, tgt.reviewer_email) END,
    pw_state_name = CASE WHEN @setOwnership = 1 THEN @pwStateName ELSE COALESCE(@pwStateName, tgt.pw_state_name) END,
    qc_stage = COALESCE(@qcStage, tgt.qc_stage),
    qc_status = COALESCE(@qcStatus, tgt.qc_status),
    last_updated_at = SYSDATETIMEOFFSET(),
    last_audit_event_at = COALESCE(@lastAuditEventAt, tgt.last_audit_event_at),
    file_modified_at = COALESCE(@fileModifiedAt, tgt.file_modified_at)
WHEN NOT MATCHED THEN INSERT
    (document_guid, document_name, document_number, folder_path, project_name, watch_root,
     extension, source_type, designer_email, reviewer_email, pw_state_name, qc_stage, qc_status,
     last_audit_event_at, file_modified_at)
VALUES
    (@docGuid, @docName, @docNumber, @folderPath, @projectName, @watchRoot,
     @extension, @sourceType, @designerEmail, @reviewerEmail, @pwStateName, @qcStage, @qcStatus,
     @lastAuditEventAt, @fileModifiedAt);
"@
        $params = @{
            docGuid          = $DocumentGuid
            docName          = $DocumentName
            docNumber        = $DocumentNumber
            folderPath       = $FolderPath
            projectName      = if ($ProjectName)      { $ProjectName }      else { $null }
            watchRoot        = if ($WatchRoot)         { $WatchRoot }         else { $null }
            extension        = if ($Extension)         { $Extension }         else { $null }
            sourceType       = if ($SourceType)        { $SourceType }        else { $null }
            designerEmail    = if ($DesignerEmail)     { $DesignerEmail }     else { $null }
            reviewerEmail    = if ($ReviewerEmail)     { $ReviewerEmail }     else { $null }
            pwStateName      = if ($PwStateName)       { $PwStateName }       else { $null }
            qcStage          = if ($QcStage)           { $QcStage }           else { $null }
            qcStatus         = if ($QcStatus)          { $QcStatus }          else { $null }
            lastAuditEventAt = if ($LastAuditEventAt)  { $LastAuditEventAt }  else { $null }
            fileModifiedAt   = if ($FileModifiedAt)    { $FileModifiedAt }    else { $null }
            setOwnership     = $setOwnership
        }
        Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
    } catch { }
}

function Update-QCSheetQcPdf {
    <#
    .SYNOPSIS
    Links a -qc.pdf to its source sheet in the sheet_index table. Fire-and-forget.
    Called after a successful QC_PREPEND job.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$SourceDocumentGuid,
        [string]$QcPdfGuid,
        [string]$QcPdfName
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    try {
        $sql = @"
UPDATE sheet_index
SET qc_pdf_guid = @qcPdfGuid,
    qc_pdf_name = @qcPdfName,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @sourceDocGuid
"@
        $params = @{
            sourceDocGuid = $SourceDocumentGuid
            qcPdfGuid     = if ($QcPdfGuid) { $QcPdfGuid } else { $null }
            qcPdfName     = if ($QcPdfName) { $QcPdfName } else { $null }
        }
        Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
    } catch { }
}

Export-ModuleMember -Function Test-QCDatabaseEnabled, Test-QCDatabaseWritesAllowed, Get-QCDatabaseConnection, Invoke-QCDatabaseQuery, Invoke-QCDatabaseNonQuery, Invoke-QCDatabaseScalar, Invoke-QCDatabaseBatch, Initialize-QCDatabaseSchema, Get-QCProcessingJobType, Write-QCJobTelemetry, Write-QCPollRunTelemetry, Write-QCNotificationTelemetry, Write-QCSheetIndex, Update-QCSheetQcPdf
