# Core.Database.psm1
# Responsibility: SQL Server connectivity and schema management for QC pipeline telemetry.
# The database is the reporting/control layer. The JSON queue remains the execution source.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Paths.psm1') -Force

function _QDB-NormalizeTelemetryPath {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $r = Normalize-QCDocumentsFolderPath -Path $Path
    if ($r.IsSuccess) { return [string]$r.Data.path }
    return $Path.Trim()
}

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

function Test-QCSheetIndexFolderPath {
    <#
    .SYNOPSIS
    Returns $true when folder_path is under a project CADD\Sheets tree (sheet_index scope).
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$FolderPath,
        [string]$RequiredFragment = 'CADD/Sheets'
    )
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $false }
    $normPath = ([string]$FolderPath).Trim() -replace '\\', '/'
    $normFrag = ([string]$RequiredFragment).Trim() -replace '\\', '/'
    if ([string]::IsNullOrWhiteSpace($normFrag)) { return $false }
    return $normPath.IndexOf($normFrag, [StringComparison]::OrdinalIgnoreCase) -ge 0
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

function _QDB-ConvertDataTableToRowHashtables {
    param([AllowNull()][System.Data.DataTable]$Table)
    if (-not $Table -or $Table.Rows.Count -eq 0) { return @() }
    $list = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($dataRow in @($Table.Rows)) {
        if ($null -eq $dataRow) { continue }
        $h = @{}
        foreach ($col in $Table.Columns) {
            $name = [string]$col.ColumnName
            $val = $dataRow[$name]
            if ($val -is [DBNull]) { $h[$name] = $null } else { $h[$name] = $val }
        }
        $list.Add($h)
    }
    return @($list)
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
# Optional connection reuse (high-volume loops)
# ---------------------------------------------------------------------------

function New-QCDatabaseSession {
    <#
    .SYNOPSIS
    Opens one SqlConnection for repeated calls; caller must Dispose().
    .DESCRIPTION
    Returns a QCResult whose Data.session has:
      - connection: open SqlConnection
      - Dispose(): closes/cleans up
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $res = Get-QCDatabaseConnection -Config $Config
    if (-not $res.IsSuccess) { return $res }
    $conn = $res.Data.connection

    $session = [pscustomobject]@{
        connection = $conn
        Dispose = {
            try { if ($this.connection -and $this.connection.State -eq 'Open') { $this.connection.Close() } } catch { }
            try { if ($this.connection) { $this.connection.Dispose() } } catch { }
        }
    }
    return New-QCSuccessResult -Code 'DB_SESSION_OK' -Message 'Database session opened.' -Data @{ session = $session }
}

function Invoke-QCDatabaseNonQueryWithConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Parameters = @{},
        [int]$CommandTimeout = 60
    )
    try {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $CommandTimeout
        foreach ($key in $Parameters.Keys) {
            $val = $Parameters[$key]
            if ($null -eq $val) { $val = [DBNull]::Value }
            [void]$cmd.Parameters.AddWithValue("@$key", $val)
        }
        $rowsAffected = $cmd.ExecuteNonQuery()
        return New-QCSuccessResult -Code 'DB_NONQUERY_OK' -Message "$rowsAffected rows affected." -Data @{ rowsAffected = $rowsAffected }
    } catch {
        return New-QCFailureResult -Code 'DB_NONQUERY_FAILED' -Message "Non-query failed: $($_.Exception.Message)" -Data @{ sql = $Sql; error = $_.Exception.Message }
    }
}

function Invoke-QCDatabaseScalarWithConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Parameters = @{},
        [int]$CommandTimeout = 60
    )
    try {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = $CommandTimeout
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
    }
}

# ---------------------------------------------------------------------------
# Schema management
# ---------------------------------------------------------------------------

function _QDB-InvokeSchemaSqlBatches {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$Sql
    )
    $batches = [regex]::Split($Sql, '^\s*GO\s*$', [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $executed = 0
    foreach ($batch in $batches) {
        $trimmed = $batch.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $batchCmd = $Connection.CreateCommand()
        $batchCmd.CommandText = $trimmed
        $batchCmd.CommandTimeout = 120
        [void]$batchCmd.ExecuteNonQuery()
        $executed++
    }
    return $executed
}

function _QDB-AcquireSchemaInitLock {
    param([Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection)
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = @"
DECLARE @rc INT;
EXEC @rc = sp_getapplock
    @Resource = 'QC_DatabaseSchemaInit',
    @LockMode = 'Exclusive',
    @LockOwner = 'Session',
    @LockTimeout = 60000;
SELECT @rc;
"@
    $cmd.CommandTimeout = 120
    $rc = $cmd.ExecuteScalar()
    if ($null -eq $rc -or [int]$rc -lt 0) {
        throw "Could not acquire schema initialization lock (sp_getapplock returned $rc)."
    }
}

function _QDB-ReleaseSchemaInitLock {
    param([Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection)
    try {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = "EXEC sp_releaseapplock @Resource = 'QC_DatabaseSchemaInit', @LockOwner = 'Session';"
        $cmd.CommandTimeout = 30
        [void]$cmd.ExecuteNonQuery()
    } catch {
        # Best-effort; connection close also releases session-owned locks.
    }
}

function Initialize-QCDatabaseSchema {
    <#
    .SYNOPSIS
    Creates all QC pipeline telemetry tables, indexes, and views idempotently.
    Tracks applied versions in schema_version table.
    Always applies additive patches so existing databases pick up new columns without a full rebuild.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'DB_DISABLED' -Message 'Database is not enabled in config.' -Data @{}
    }

    $targetVersion = '1.9.0'
    $schemaV1 = _QDB-GetSchemaV1
    $schemaV1_1 = _QDB-GetSchemaV1dot1
    $schemaV1_2 = _QDB-GetSchemaV1dot2
    $schemaV1_3 = _QDB-GetSchemaV1dot3
    $schemaV1_4 = _QDB-GetSchemaV1dot4
    $schemaV1_5 = _QDB-GetSchemaV1dot5
    $schemaV1_6 = _QDB-GetSchemaV1dot6
    $schemaSql = $schemaV1 + [Environment]::NewLine + $schemaV1_1 + [Environment]::NewLine + $schemaV1_2 + [Environment]::NewLine + $schemaV1_3 + [Environment]::NewLine + $schemaV1_4 + [Environment]::NewLine + $schemaV1_5 + [Environment]::NewLine + $schemaV1_6
    $patchSql = (_QDB-GetSchemaV1dot3Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot4Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot5Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot6Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot7Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot8Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot9Additive) + [Environment]::NewLine + (_QDB-GetProcessingJobsAdditive)

    $connRes = Get-QCDatabaseConnection -Config $Config
    if (-not $connRes.IsSuccess) { return $connRes }
    $conn = $connRes.Data.connection
    $lockHeld = $false

    try {
        _QDB-AcquireSchemaInitLock -Connection $conn
        $lockHeld = $true

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "IF OBJECT_ID('dbo.schema_version', 'U') IS NOT NULL SELECT MAX(version) FROM schema_version ELSE SELECT NULL"
        $currentVersion = $cmd.ExecuteScalar()
        if ($currentVersion -is [DBNull]) { $currentVersion = $null }
        if ($currentVersion) { $currentVersion = [string]$currentVersion }

        $patchCount = _QDB-InvokeSchemaSqlBatches -Connection $conn -Sql $patchSql

        if ($currentVersion -eq $targetVersion) {
            return New-QCSuccessResult -Code 'DB_SCHEMA_CURRENT' -Message "Schema already at version $targetVersion (additive patches: $patchCount)." -Data @{ version = $targetVersion; patchCount = $patchCount; previousVersion = $currentVersion }
        }

        $bootstrapCount = 0
        if ($null -eq $currentVersion) {
            $bootstrapCount = _QDB-InvokeSchemaSqlBatches -Connection $conn -Sql $schemaSql
        }

        $insertCmd = $conn.CreateCommand()
        $insertCmd.CommandText = @"
IF NOT EXISTS (SELECT 1 FROM schema_version WHERE version = @version)
INSERT INTO schema_version (version, description) VALUES (@version, @desc)
"@
        [void]$insertCmd.Parameters.AddWithValue("@version", $targetVersion)
        [void]$insertCmd.Parameters.AddWithValue("@desc", "QC telemetry schema through transition_events.changed_by columns")
        [void]$insertCmd.ExecuteNonQuery()

        if ($null -eq $currentVersion) {
            return New-QCSuccessResult -Code 'DB_SCHEMA_INITIALIZED' -Message "Schema initialized to version $targetVersion ($bootstrapCount bootstrap batches, $patchCount patches)." -Data @{ version = $targetVersion; batchCount = $bootstrapCount; patchCount = $patchCount; previousVersion = $null }
        }

        return New-QCSuccessResult -Code 'DB_SCHEMA_UPGRADED' -Message "Schema upgraded from $currentVersion to $targetVersion ($patchCount patches)." -Data @{ version = $targetVersion; patchCount = $patchCount; previousVersion = $currentVersion }
    } catch {
        return New-QCFailureResult -Code 'DB_SCHEMA_FAILED' -Message "Schema initialization failed: $($_.Exception.Message)" -Data @{ error = $_.Exception.Message }
    } finally {
        if ($lockHeld -and $conn -and $conn.State -eq 'Open') {
            _QDB-ReleaseSchemaInitLock -Connection $conn
        }
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
    is_reconciliation   BIT NOT NULL DEFAULT 0,
    watcher_name        NVARCHAR(120),
    service_name        NVARCHAR(120),
    pass_number         INT,
    run_mode            NVARCHAR(40),
    run_status          NVARCHAR(20),
    total_duration_seconds DECIMAL(18,3),
    audit_query_duration_seconds DECIMAL(18,3),
    reconciliation_duration_seconds DECIMAL(18,3),
    trigger_eval_duration_seconds DECIMAL(18,3),
    dedupe_duration_seconds DECIMAL(18,3),
    queue_write_duration_seconds DECIMAL(18,3),
    database_write_duration_seconds DECIMAL(18,3),
    cleanup_duration_seconds DECIMAL(18,3),
    sleep_throttle_duration_seconds DECIMAL(18,3),
    candidate_documents_evaluated INT NOT NULL DEFAULT 0,
    trigger_matches INT NOT NULL DEFAULT 0,
    jobs_skipped_dedupe INT NOT NULL DEFAULT 0,
    warning_count INT NOT NULL DEFAULT 0,
    error_count INT NOT NULL DEFAULT 0,
    reconciliation_reason NVARCHAR(200),
    reconciliation_trigger_source NVARCHAR(80),
    downtime_seconds INT,
    audit_gap_detected BIT NOT NULL DEFAULT 0,
    watcher_phase NVARCHAR(80),
    throttle_wait_seconds DECIMAL(18,3),
    queue_depth_snapshot INT,
    pass_number_source  NVARCHAR(30)
);


IF COL_LENGTH('dbo.poll_runs','watcher_name') IS NULL ALTER TABLE dbo.poll_runs ADD watcher_name NVARCHAR(120) NULL;
IF COL_LENGTH('dbo.poll_runs','service_name') IS NULL ALTER TABLE dbo.poll_runs ADD service_name NVARCHAR(120) NULL;
IF COL_LENGTH('dbo.poll_runs','pass_number') IS NULL ALTER TABLE dbo.poll_runs ADD pass_number INT NULL;
IF COL_LENGTH('dbo.poll_runs','run_mode') IS NULL ALTER TABLE dbo.poll_runs ADD run_mode NVARCHAR(40) NULL;
IF COL_LENGTH('dbo.poll_runs','run_status') IS NULL ALTER TABLE dbo.poll_runs ADD run_status NVARCHAR(20) NULL;
IF COL_LENGTH('dbo.poll_runs','total_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD total_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','audit_query_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD audit_query_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','reconciliation_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD reconciliation_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','trigger_eval_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD trigger_eval_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','dedupe_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD dedupe_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','queue_write_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD queue_write_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','database_write_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD database_write_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','cleanup_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD cleanup_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','sleep_throttle_duration_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD sleep_throttle_duration_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','candidate_documents_evaluated') IS NULL ALTER TABLE dbo.poll_runs ADD candidate_documents_evaluated INT NOT NULL CONSTRAINT DF_poll_runs_candidate_documents_evaluated DEFAULT 0;
IF COL_LENGTH('dbo.poll_runs','trigger_matches') IS NULL ALTER TABLE dbo.poll_runs ADD trigger_matches INT NOT NULL CONSTRAINT DF_poll_runs_trigger_matches DEFAULT 0;
IF COL_LENGTH('dbo.poll_runs','jobs_skipped_dedupe') IS NULL ALTER TABLE dbo.poll_runs ADD jobs_skipped_dedupe INT NOT NULL CONSTRAINT DF_poll_runs_jobs_skipped_dedupe DEFAULT 0;
IF COL_LENGTH('dbo.poll_runs','warning_count') IS NULL ALTER TABLE dbo.poll_runs ADD warning_count INT NOT NULL CONSTRAINT DF_poll_runs_warning_count DEFAULT 0;
IF COL_LENGTH('dbo.poll_runs','error_count') IS NULL ALTER TABLE dbo.poll_runs ADD error_count INT NOT NULL CONSTRAINT DF_poll_runs_error_count DEFAULT 0;
IF COL_LENGTH('dbo.poll_runs','reconciliation_reason') IS NULL ALTER TABLE dbo.poll_runs ADD reconciliation_reason NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.poll_runs','pass_number_source') IS NULL ALTER TABLE dbo.poll_runs ADD pass_number_source NVARCHAR(30) NULL;


IF COL_LENGTH('dbo.poll_runs','reconciliation_trigger_source') IS NULL ALTER TABLE dbo.poll_runs ADD reconciliation_trigger_source NVARCHAR(80) NULL;
IF COL_LENGTH('dbo.poll_runs','downtime_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD downtime_seconds INT NULL;
IF COL_LENGTH('dbo.poll_runs','audit_gap_detected') IS NULL ALTER TABLE dbo.poll_runs ADD audit_gap_detected BIT NOT NULL CONSTRAINT DF_poll_runs_audit_gap_detected DEFAULT 0;
IF COL_LENGTH('dbo.poll_runs','watcher_phase') IS NULL ALTER TABLE dbo.poll_runs ADD watcher_phase NVARCHAR(80) NULL;
IF COL_LENGTH('dbo.poll_runs','throttle_wait_seconds') IS NULL ALTER TABLE dbo.poll_runs ADD throttle_wait_seconds DECIMAL(18,3) NULL;
IF COL_LENGTH('dbo.poll_runs','queue_depth_snapshot') IS NULL ALTER TABLE dbo.poll_runs ADD queue_depth_snapshot INT NULL;

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
    qc_review_type          NVARCHAR(100) NULL,
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

function _QDB-GetSchemaV1dot3 {
    return @'

GO

-- audit_events: enforce natural key uniqueness for fast set-based ingestion
IF OBJECT_ID('dbo.audit_events', 'U') IS NOT NULL
BEGIN
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_audit_events_natural_key')
        CREATE UNIQUE INDEX UX_audit_events_natural_key ON audit_events(pw_acttime, pw_action, pw_objguid)
        WHERE pw_objguid IS NOT NULL;
END

'@
}

function _QDB-GetSchemaV1dot3Additive {
    return @'
GO
IF OBJECT_ID('dbo.audit_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_audit_events_natural_key')
    CREATE UNIQUE INDEX UX_audit_events_natural_key ON audit_events(pw_acttime, pw_action, pw_objguid)
    WHERE pw_objguid IS NOT NULL;
'@
}

function _QDB-GetSchemaV1dot4 {
    return @'

GO

-- pw_users: map ProjectWise o_userno to login name and email
IF OBJECT_ID('dbo.pw_users', 'U') IS NULL
CREATE TABLE pw_users (
    pw_userno       INT NOT NULL PRIMARY KEY,
    pw_username     NVARCHAR(128) NULL,
    pw_user_email   NVARCHAR(320) NULL,
    display_name    NVARCHAR(256) NULL,
    first_seen_at   DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    last_synced_at  DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);

IF OBJECT_ID('dbo.pw_users', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_users_username')
    CREATE INDEX IX_pw_users_username ON pw_users(pw_username) WHERE pw_username IS NOT NULL;
IF OBJECT_ID('dbo.pw_users', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_users_email')
    CREATE INDEX IX_pw_users_email ON pw_users(pw_user_email) WHERE pw_user_email IS NOT NULL;

GO

IF OBJECT_ID('dbo.v_audit_events_with_user', 'V') IS NULL
EXEC('CREATE VIEW v_audit_events_with_user AS
SELECT ae.id, ae.captured_at, ae.poll_run_id, ae.pw_acttime, ae.pw_action, ae.pw_action_name,
       ae.pw_objguid, ae.pw_parentguid, ae.pw_userno, pu.pw_username, pu.pw_user_email, pu.display_name,
       ae.pw_itemname, ae.pw_itemdesc, ae.resolved_folder, ae.candidate_type, ae.processed
FROM audit_events ae
LEFT JOIN pw_users pu ON pu.pw_userno = ae.pw_userno');

'@
}

function _QDB-GetSchemaV1dot5 {
    return @'

GO

IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.sheet_index', 'checker_email') IS NULL
        ALTER TABLE sheet_index ADD checker_email NVARCHAR(200) NULL;
    IF COL_LENGTH('dbo.sheet_index', 'qc_review_type') IS NULL
        ALTER TABLE sheet_index ADD qc_review_type NVARCHAR(100) NULL;
    IF COL_LENGTH('dbo.sheet_index', 'qc_assigned_to') IS NULL
        ALTER TABLE sheet_index ADD qc_assigned_to NVARCHAR(200) NULL;
END

GO

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
    s.checker_email,
    s.qc_review_type,
    s.qc_assigned_to,
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

function _QDB-GetProcessingJobsAdditive {
    <#
    Ensures processing_jobs exists on databases that were initialized before this table
    was part of the bootstrap script (schema_version already set skips full bootstrap).
    #>
    return @'
GO
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
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_status')
    CREATE INDEX idx_jobs_status ON processing_jobs(status);
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_type')
    CREATE INDEX idx_jobs_type ON processing_jobs(job_type);
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_created')
    CREATE INDEX idx_jobs_created ON processing_jobs(created_at);
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_jobs_folder')
    CREATE INDEX idx_jobs_folder ON processing_jobs(source_folder);
GO
IF OBJECT_ID('dbo.v_job_summary', 'V') IS NULL AND OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL
EXEC('CREATE VIEW v_job_summary AS
SELECT
    job_type,
    status,
    COUNT(*) AS job_count,
    AVG(duration_ms) AS avg_duration_ms,
    MAX(completed_at) AS last_completed
FROM processing_jobs
GROUP BY job_type, status');
'@
}

function _QDB-GetSchemaV1dot5Additive {
    return @'
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_index', 'checker_email') IS NULL
    ALTER TABLE sheet_index ADD checker_email NVARCHAR(200) NULL;
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_index', 'qc_review_type') IS NULL
    ALTER TABLE sheet_index ADD qc_review_type NVARCHAR(100) NULL;
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_index', 'qc_assigned_to') IS NULL
    ALTER TABLE sheet_index ADD qc_assigned_to NVARCHAR(200) NULL;
GO
IF OBJECT_ID('dbo.v_sheet_status', 'V') IS NOT NULL
    DROP VIEW dbo.v_sheet_status;
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL
EXEC('CREATE VIEW dbo.v_sheet_status AS
SELECT
    s.document_name,
    s.folder_path,
    s.project_name,
    s.extension,
    s.designer_email,
    s.reviewer_email,
    s.checker_email,
    s.qc_review_type,
    s.qc_assigned_to,
    s.pw_state_name,
    s.qc_stage,
    s.qc_status,
    s.qc_pdf_name,
    CASE WHEN s.qc_pdf_guid IS NOT NULL THEN 1 ELSE 0 END AS has_qc_pdf,
    s.last_updated_at,
    s.file_modified_at,
    s.watch_root
FROM sheet_index s');
'@
}

function _QDB-GetSchemaV1dot6 {
    return @'

GO

-- watcher_state: durable watcher cursor and operational flags (DB source of truth for audit watermark)
IF OBJECT_ID('dbo.watcher_state', 'U') IS NULL
CREATE TABLE watcher_state (
    state_key       NVARCHAR(100) NOT NULL PRIMARY KEY,
    state_value     NVARCHAR(500) NOT NULL,
    updated_at      DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);

GO

-- pw_document_cache: TTL metadata cache to minimize repeated PW GUID lookups
IF OBJECT_ID('dbo.pw_document_cache', 'U') IS NULL
CREATE TABLE pw_document_cache (
    document_guid       NVARCHAR(50) NOT NULL PRIMARY KEY,
    folder_path         NVARCHAR(1000) NULL,
    description         NVARCHAR(1000) NULL,
    workflow_state      NVARCHAR(200) NULL,
    resolve_failed      BIT NOT NULL DEFAULT 0,
    last_audit_action   NVARCHAR(100) NULL,
    cached_at           DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    expires_at          DATETIMEOFFSET(3) NOT NULL
);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_document_cache_expires')
    CREATE INDEX IX_pw_document_cache_expires ON pw_document_cache(expires_at);

GO

IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.processing_jobs', 'last_heartbeat_at') IS NULL
        ALTER TABLE processing_jobs ADD last_heartbeat_at DATETIMEOFFSET(3) NULL;
    IF COL_LENGTH('dbo.processing_jobs', 'recovery_count') IS NULL
        ALTER TABLE processing_jobs ADD recovery_count INT NOT NULL DEFAULT 0;
    IF COL_LENGTH('dbo.processing_jobs', 'recovery_reason') IS NULL
        ALTER TABLE processing_jobs ADD recovery_reason NVARCHAR(200) NULL;
    IF COL_LENGTH('dbo.processing_jobs', 'checkpoint') IS NULL
        ALTER TABLE processing_jobs ADD [checkpoint] NVARCHAR(100) NULL;
    IF COL_LENGTH('dbo.processing_jobs', 'checkpoint_data') IS NULL
        ALTER TABLE processing_jobs ADD [checkpoint_data] NVARCHAR(MAX) NULL;
END

'@
}

function _QDB-GetSchemaV1dot6Additive {
    return @'
GO
IF OBJECT_ID('dbo.watcher_state', 'U') IS NULL
CREATE TABLE watcher_state (
    state_key       NVARCHAR(100) NOT NULL PRIMARY KEY,
    state_value     NVARCHAR(500) NOT NULL,
    updated_at      DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);
IF OBJECT_ID('dbo.pw_document_cache', 'U') IS NULL
CREATE TABLE pw_document_cache (
    document_guid       NVARCHAR(50) NOT NULL PRIMARY KEY,
    folder_path         NVARCHAR(1000) NULL,
    description         NVARCHAR(1000) NULL,
    workflow_state      NVARCHAR(200) NULL,
    resolve_failed      BIT NOT NULL DEFAULT 0,
    last_audit_action   NVARCHAR(100) NULL,
    cached_at           DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    expires_at          DATETIMEOFFSET(3) NOT NULL
);
IF OBJECT_ID('dbo.pw_document_cache', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_document_cache_expires')
    CREATE INDEX IX_pw_document_cache_expires ON pw_document_cache(expires_at);
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'last_heartbeat_at') IS NULL
    ALTER TABLE processing_jobs ADD last_heartbeat_at DATETIMEOFFSET(3) NULL;
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'recovery_count') IS NULL
    ALTER TABLE processing_jobs ADD recovery_count INT NOT NULL DEFAULT 0;
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'recovery_reason') IS NULL
    ALTER TABLE processing_jobs ADD recovery_reason NVARCHAR(200) NULL;
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'checkpoint') IS NULL
    ALTER TABLE processing_jobs ADD [checkpoint] NVARCHAR(100) NULL;
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'checkpoint_data') IS NULL
    ALTER TABLE processing_jobs ADD [checkpoint_data] NVARCHAR(MAX) NULL;
'@
}

function _QDB-GetSchemaV1dot7Additive {
    return @'
GO
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'qc_review_type') IS NULL
    ALTER TABLE qc_workflow_events ADD qc_review_type NVARCHAR(100) NULL;
'@
}

function _QDB-GetSchemaV1dot8Additive {
    return @'
GO
IF OBJECT_ID('dbo.transition_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.transition_events', 'changed_by_user') IS NULL
    ALTER TABLE transition_events ADD changed_by_user INT NULL;
IF OBJECT_ID('dbo.transition_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.transition_events', 'changed_by_username') IS NULL
    ALTER TABLE transition_events ADD changed_by_username NVARCHAR(128) NULL;
IF OBJECT_ID('dbo.document_state_history', 'U') IS NOT NULL AND COL_LENGTH('dbo.document_state_history', 'changed_by_username') IS NULL
    ALTER TABLE document_state_history ADD changed_by_username NVARCHAR(128) NULL;
'@
}

function _QDB-GetSchemaV1dot9Additive {
    return @'
GO
IF OBJECT_ID('dbo.pw_folder_cache', 'U') IS NULL
CREATE TABLE pw_folder_cache (
    folder_guid       NVARCHAR(50) NOT NULL PRIMARY KEY,
    folder_path       NVARCHAR(1000) NULL,
    watch_root        NVARCHAR(1000) NULL,
    resolve_failed    BIT NOT NULL DEFAULT 0,
    cached_at         DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    expires_at        DATETIMEOFFSET(3) NOT NULL
);
IF OBJECT_ID('dbo.pw_folder_cache', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_folder_cache_expires')
    CREATE INDEX IX_pw_folder_cache_expires ON pw_folder_cache(expires_at);
IF OBJECT_ID('dbo.pw_folder_cache', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_folder_cache_path')
    CREATE INDEX IX_pw_folder_cache_path ON pw_folder_cache(folder_path) WHERE folder_path IS NOT NULL;
'@
}

function _QDB-GetSchemaV1dot4Additive {
    return @'
GO
IF OBJECT_ID('dbo.pw_users', 'U') IS NULL
CREATE TABLE pw_users (
    pw_userno       INT NOT NULL PRIMARY KEY,
    pw_username     NVARCHAR(128) NULL,
    pw_user_email   NVARCHAR(320) NULL,
    display_name    NVARCHAR(256) NULL,
    first_seen_at   DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET(),
    last_synced_at  DATETIMEOFFSET(3) NOT NULL DEFAULT SYSDATETIMEOFFSET()
);
IF OBJECT_ID('dbo.pw_users', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_users_username')
    CREATE INDEX IX_pw_users_username ON pw_users(pw_username) WHERE pw_username IS NOT NULL;
IF OBJECT_ID('dbo.pw_users', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_pw_users_email')
    CREATE INDEX IX_pw_users_email ON pw_users(pw_user_email) WHERE pw_user_email IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_audit_events_with_user', 'V') IS NULL
EXEC('CREATE VIEW v_audit_events_with_user AS
SELECT ae.id, ae.captured_at, ae.poll_run_id, ae.pw_acttime, ae.pw_action, ae.pw_action_name,
       ae.pw_objguid, ae.pw_parentguid, ae.pw_userno, pu.pw_username, pu.pw_user_email, pu.display_name,
       ae.pw_itemname, ae.pw_itemdesc, ae.resolved_folder, ae.candidate_type, ae.processed
FROM audit_events ae
LEFT JOIN pw_users pu ON pu.pw_userno = ae.pw_userno');
'@
}

function Get-QCPWUnresolvedUserNumbers {
    <#
    .SYNOPSIS
    Returns pw_userno values present in audit_events but not yet in pw_users.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [int]$MaxCount = 100
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'DB_DISABLED' -Message 'Database is not enabled.' -Data @{}
    }
    $sql = @"
SELECT DISTINCT TOP (@maxCount) ae.pw_userno
FROM audit_events ae
WHERE ae.pw_userno IS NOT NULL AND ae.pw_userno > 0
  AND NOT EXISTS (SELECT 1 FROM pw_users pu WHERE pu.pw_userno = ae.pw_userno)
ORDER BY ae.pw_userno;
"@
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters @{ maxCount = $MaxCount }
    if (-not $res.IsSuccess) { return $res }
    $numbers = @()
    $table = $res.Data.table
    if ($table) {
        foreach ($row in $table.Rows) {
            if ($null -eq $row -or $row.IsNull('pw_userno')) { continue }
            try { $numbers += [int]$row['pw_userno'] } catch { }
        }
    }
    return New-QCSuccessResult -Code 'PW_USER_NUMBERS_OK' -Message "Found $($numbers.Count) unresolved user number(s)." -Data @{ numbers = $numbers }
}

function Get-QCPWUserIdentity {
    <#
    .SYNOPSIS
    Returns pw_users identity for a ProjectWise user number (email preferred for display).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][int]$UserNumber
    )
    if ($UserNumber -le 0) {
        return New-QCSuccessResult -Code 'PW_USER_IDENTITY_INVALID' -Message 'User number must be positive.' -Data @{ identity = $null }
    }
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_USER_IDENTITY_SKIPPED' -Message 'Database is disabled.' -Data @{ identity = $null }
    }
    $sql = @"
SELECT pw_userno, pw_username, pw_user_email, display_name
FROM pw_users
WHERE pw_userno = @userno
"@
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters @{ userno = $UserNumber }
    if (-not $res.IsSuccess) { return $res }
    $table = $res.Data.table
    if (-not $table -or $table.Rows.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_USER_IDENTITY_NOT_FOUND' -Message 'User not in pw_users.' -Data @{ identity = $null; pwUserno = $UserNumber }
    }
    $row = $table.Rows[0]
    $identity = @{
        pw_userno = $UserNumber
        pw_username = if ($row.IsNull('pw_username')) { '' } else { [string]$row['pw_username'] }
        pw_user_email = if ($row.IsNull('pw_user_email')) { '' } else { [string]$row['pw_user_email'] }
        display_name = if ($row.IsNull('display_name')) { '' } else { [string]$row['display_name'] }
    }
    return New-QCSuccessResult -Code 'PW_USER_IDENTITY_OK' -Message 'User identity resolved from pw_users.' -Data @{ identity = $identity; pwUserno = $UserNumber }
}

function Write-QCPWUserDirectory {
    <#
    .SYNOPSIS
    Upserts ProjectWise user identity rows into dbo.pw_users.
  Each item in -Users should be a hashtable with pw_userno and optional pw_username, pw_user_email, display_name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Users
    )
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_USER_WRITE_SKIPPED' -Message 'Database writes not allowed.' -Data @{ rowsAffected = 0 }
    }
    $valid = @($Users | Where-Object { $_ -and $_.pw_userno -gt 0 })
    if ($valid.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_USER_WRITE_NONE' -Message 'No valid user rows.' -Data @{ rowsAffected = 0 }
    }

    $chunkSize = 50
    $totalAffected = 0
    for ($i = 0; $i -lt $valid.Count; $i += $chunkSize) {
        $chunk = @($valid[$i..[Math]::Min($i + $chunkSize - 1, $valid.Count - 1)])
        $valuesSql = New-Object System.Text.StringBuilder
        $params = @{}
        for ($r = 0; $r -lt $chunk.Count; $r++) {
            $row = $chunk[$r]
            if ($r -gt 0) { [void]$valuesSql.AppendLine(',') }
            [void]$valuesSql.Append(("(@userno{0},@username{0},@email{0},@display{0})" -f $r))
            $params["userno$r"] = [int]$row.pw_userno
            $params["username$r"] = if ($row.pw_username) { [string]$row.pw_username } else { $null }
            $params["email$r"] = if ($row.pw_user_email) { [string]$row.pw_user_email } else { $null }
            $params["display$r"] = if ($row.display_name) { [string]$row.display_name } else { $null }
        }

        $sql = @"
MERGE pw_users AS tgt
USING (VALUES
$($valuesSql.ToString())
) AS src(pw_userno, pw_username, pw_user_email, display_name)
ON tgt.pw_userno = src.pw_userno
WHEN MATCHED THEN UPDATE SET
    pw_username = COALESCE(src.pw_username, tgt.pw_username),
    pw_user_email = COALESCE(src.pw_user_email, tgt.pw_user_email),
    display_name = COALESCE(src.display_name, tgt.display_name),
    last_synced_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (pw_userno, pw_username, pw_user_email, display_name)
VALUES (src.pw_userno, src.pw_username, src.pw_user_email, src.display_name);
"@
        $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $res.IsSuccess) { return $res }
        if ($res.Data.rowsAffected) { $totalAffected += [int]$res.Data.rowsAffected }
    }

    return New-QCSuccessResult -Code 'PW_USER_WRITE_OK' -Message "Upserted user directory ($totalAffected row(s) affected)." -Data @{ rowsAffected = $totalAffected; userCount = $valid.Count }
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

function _QDB-GetMaxRowsForSqlParameters {
    <#
    SQL Server caps parameters at 2100 per batch. Used to size multi-row VALUES inserts.
    #>
    param(
        [Parameter(Mandatory)][int]$ParametersPerRow,
        [int]$Reserved = 10,
        [int]$ServerMax = 2100
    )
    $ppr = [Math]::Max(1, $ParametersPerRow)
    return [Math]::Max(1, [int][Math]::Floor(($ServerMax - $Reserved) / $ppr))
}

function _QDB-PrepareAuditEventRowsForInsert {
    param([array]$Rows)
    $prepared = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $skippedNoGuid = 0
    $skippedDupInBatch = 0
    foreach ($row in $Rows) {
        $g = [string]$row.objguid
        if ([string]::IsNullOrWhiteSpace($g)) { $skippedNoGuid++; continue }
        $key = ([string]$row.acttime) + '|' + [string]$row.action + '|' + $g.Trim()
        if (-not $seen.Add($key)) { $skippedDupInBatch++; continue }
        $prepared.Add($row)
    }
    return @{
        rows              = @($prepared)
        skippedNoGuid     = $skippedNoGuid
        skippedDupInBatch = $skippedDupInBatch
    }
}

function Write-QCAuditEventRows {
    <#
    .SYNOPSIS
    Batch-insert PW audit trail rows into audit_events (deduped by natural key).
    Rows without a document GUID are skipped (USER_LOGIN etc. use empty guid and break UX_audit_events_natural_key).
    .OUTPUTS
    QCResult with Data: @{ written; skipped; failed; skippedNoGuid; skippedDupInBatch; chunks; lastError }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Rows,
        [int]$ChunkSize = 150
    )

    $auditParamsPerRow = 13
    $maxChunk = _QDB-GetMaxRowsForSqlParameters -ParametersPerRow $auditParamsPerRow
    if ($ChunkSize -lt 1 -or $ChunkSize -gt $maxChunk) { $ChunkSize = $maxChunk }

    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'AUDIT_EVENTS_SKIPPED' -Message 'Database disabled.' -Data @{ written = 0; skipped = $Rows.Count; failed = 0; chunks = 0 }
    }
    if (-not $Rows -or $Rows.Count -eq 0) {
        return New-QCSuccessResult -Code 'AUDIT_EVENTS_NONE' -Message 'No audit rows to write.' -Data @{ written = 0; skipped = 0; failed = 0; chunks = 0 }
    }

    $prep = _QDB-PrepareAuditEventRowsForInsert -Rows $Rows
    $insertRows = @($prep.rows)
    $skipped = [int]$prep.skippedNoGuid + [int]$prep.skippedDupInBatch
    if ($insertRows.Count -eq 0) {
        return New-QCSuccessResult -Code 'AUDIT_EVENTS_NONE_INSERTABLE' -Message 'No audit rows with document GUID to insert.' -Data @{
            written = 0; skipped = $skipped; skippedNoGuid = [int]$prep.skippedNoGuid
            skippedDupInBatch = [int]$prep.skippedDupInBatch; failed = 0; chunks = 0; lastError = $null
        }
    }

    $written = 0
    $failed = 0
    $chunks = 0
    $lastError = $null
    for ($i = 0; $i -lt $insertRows.Count; $i += $ChunkSize) {
        $chunk = @($insertRows[$i..[Math]::Min($i + $ChunkSize - 1, $insertRows.Count - 1)])
        $valuesSql = New-Object System.Text.StringBuilder
        $params = @{}
        for ($r = 0; $r -lt $chunk.Count; $r++) {
            $row = $chunk[$r]
            if ($r -gt 0) { [void]$valuesSql.AppendLine(',') }
            [void]$valuesSql.Append(("(@acttime{0},@action{0},@actionName{0},@objtype{0},@objno{0},@objguid{0},@parentguid{0},@userno{0},@itemname{0},@itemdesc{0},@textparam{0},@folder{0},@candidateType{0})" -f $r))
            foreach ($k in @('acttime','action','actionName','objtype','objno','objguid','parentguid','userno','itemname','itemdesc','textparam','folder','candidateType')) {
                $params[("$k$r")] = $row[$k]
            }
        }

        $insertSql = @"
INSERT INTO audit_events
    (pw_acttime, pw_action, pw_action_name, pw_objtype, pw_objno, pw_objguid, pw_parentguid, pw_userno, pw_itemname, pw_itemdesc, pw_textparam, resolved_folder, candidate_type)
SELECT v.pw_acttime, v.pw_action, v.pw_action_name, v.pw_objtype, v.pw_objno, v.pw_objguid, v.pw_parentguid, v.pw_userno, v.pw_itemname, v.pw_itemdesc, v.pw_textparam, v.resolved_folder, v.candidate_type
FROM (VALUES
$($valuesSql.ToString())
) AS v(pw_acttime, pw_action, pw_action_name, pw_objtype, pw_objno, pw_objguid, pw_parentguid, pw_userno, pw_itemname, pw_itemdesc, pw_textparam, resolved_folder, candidate_type)
WHERE NULLIF(LTRIM(RTRIM(v.pw_objguid)), '') IS NOT NULL
  AND NOT EXISTS (
    SELECT 1 FROM audit_events ae
    WHERE ae.pw_acttime = v.pw_acttime AND ae.pw_action = v.pw_action AND ae.pw_objguid = v.pw_objguid
  );
"@
        $chunks++
        try {
            $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql $insertSql -Parameters $params
            if ($res.IsSuccess) {
                $written += [int]$res.Data.rowsAffected
                $skipped += ($chunk.Count - [int]$res.Data.rowsAffected)
            } else {
                $skipped += $chunk.Count
                $failed += $chunk.Count
                $lastError = [string]$res.Message
                if ($res.Data -and $res.Data.error) { $lastError = "$lastError ($($res.Data.error))" }
            }
        } catch {
            $skipped += $chunk.Count
            $failed += $chunk.Count
            $lastError = [string]$_.Exception.Message
        }
    }

    return New-QCSuccessResult -Code 'AUDIT_EVENTS_WRITTEN' -Message "Audit events: $written inserted, $skipped skipped/duplicate." -Data @{
        written           = $written
        skipped           = $skipped
        failed            = $failed
        skippedNoGuid     = [int]$prep.skippedNoGuid
        skippedDupInBatch = [int]$prep.skippedDupInBatch
        chunks            = $chunks
        lastError         = $lastError
    }
}

# ---------------------------------------------------------------------------
# Fire-and-forget telemetry writers
# These silently no-op when the database is disabled or unreachable.
# Pipeline execution must NEVER fail because telemetry fails.
# ---------------------------------------------------------------------------

function _QDB-SafeWrite {
    param([hashtable]$Config, [string]$Sql, [hashtable]$Parameters)
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'POLL_RUN_TELEMETRY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
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
        'QC_STATE'               = 'QC_STATE'
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

function _QDB-TruncateTelemetryPayload {
    param(
        [string]$Text,
        [int]$MaxLength = 32000
    )
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    if ($Text.Length -le $MaxLength) { return $Text }
    return $Text.Substring(0, $MaxLength)
}

function New-QCStateChangeJobId {
    <#
    .SYNOPSIS
    Builds a stable processing_jobs.job_id for an automation state write (separate from queue prepend jobs).
    #>
    [CmdletBinding()]
    param(
        [string]$ParentJobId = '',
        [string]$DocumentGuid = '',
        [string]$PreviousState = '',
        [string]$CurrentState = '',
        [string]$Operation = 'state'
    )
    if (-not [string]::IsNullOrWhiteSpace($ParentJobId)) {
        return ([string]$ParentJobId).Trim() + '|state'
    }
    $g = if ([string]::IsNullOrWhiteSpace($DocumentGuid)) {
        [guid]::NewGuid().ToString('n')
    } else {
        $DocumentGuid.Trim().ToLowerInvariant()
    }
    $from = if ($PreviousState) { $PreviousState.Trim().Replace(' ', '_') } else { '_' }
    $to = if ($CurrentState) { $CurrentState.Trim().Replace(' ', '_') } else { '_' }
    return "qc-$Operation-$g-$from-$to"
}

function Write-QCStateChangeJobTelemetry {
    <#
    .SYNOPSIS
    Records a workflow state change performed by QC automation in processing_jobs as job_type QC_STATE.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$PreviousState,
        [Parameter(Mandatory)][string]$CurrentState,
        [string]$JobId = '',
        [string]$ParentJobId = '',
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$SourceFolder = '',
        [string]$TriggerSource = 'automation',
        [string]$Operation = 'state',
        [string]$Status = 'succeeded',
        [string]$ErrorMessage = '',
        [int]$DurationMs = 0,
        [string]$QcReviewType = ''
    )

    $prev = if ($PreviousState) { $PreviousState.Trim() } else { '' }
    $curr = if ($CurrentState) { $CurrentState.Trim() } else { '' }
    if ($prev -eq $curr) { return }

    if ([string]::IsNullOrWhiteSpace($JobId)) {
        $JobId = New-QCStateChangeJobId -ParentJobId $ParentJobId -DocumentGuid $DocumentGuid `
            -PreviousState $prev -CurrentState $curr -Operation $Operation
    }

    $sourcePath = $null
    if ($SourceFolder -and $DocumentName) {
        $sourcePath = ($SourceFolder.TrimEnd('\') + '\' + $DocumentName)
    }

    $resultObj = @{
        previousState = $prev
        currentState  = $curr
        documentGuid  = $DocumentGuid
        documentName  = $DocumentName
        operation     = $Operation
    }
    if (-not [string]::IsNullOrWhiteSpace($QcReviewType)) { $resultObj['qc_review_type'] = $QcReviewType }
    $resultJson = $null
    try { $resultJson = ($resultObj | ConvertTo-Json -Compress) } catch { }

    return Write-QCJobTelemetry -Config $Config -JobId $JobId -JobType 'QC_STATE' -Status $Status `
        -SourcePath $sourcePath -SourceFolder $SourceFolder -TriggerSource $TriggerSource `
        -DurationMs $(if ($DurationMs -gt 0) { $DurationMs } else { $null }) `
        -ErrorMessage $(if ($ErrorMessage) { $ErrorMessage } else { $null }) `
        -ResultData $resultJson
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
        [string]$StartedAtUtc,
        [int]$AttemptCount = 0,
        [Nullable[int]]$DurationMs,
        [string]$ErrorCode,
        [string]$ErrorMessage,
        [string]$ResultData
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'JOB_TELEMETRY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false; reason = 'database_disabled' }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'JOB_TELEMETRY_SKIPPED' -Message 'Database writes blocked (dry-run).' -Data @{ written = $false; reason = 'dry_run' }
    }
    $SourceFolder = _QDB-NormalizeTelemetryPath -Path $SourceFolder
    $SourcePath = _QDB-NormalizeTelemetryPath -Path $SourcePath
    $telemetryJobType = Get-QCProcessingJobType -QueueJobType $JobType -Config $Config
    $resultPayload = _QDB-TruncateTelemetryPayload -Text $ResultData
    $startedAt = $null
    if (-not [string]::IsNullOrWhiteSpace($StartedAtUtc)) {
        try { $startedAt = [DateTimeOffset]::Parse($StartedAtUtc) } catch { $startedAt = $null }
    }
    try {
        $sql = @"
MERGE processing_jobs AS tgt
USING (SELECT @jobId AS job_id) AS src ON tgt.job_id = src.job_id
WHEN MATCHED THEN UPDATE SET
    status = @status,
    started_at = COALESCE(@startedAt, tgt.started_at),
    completed_at = CASE WHEN @status IN ('succeeded','failed') THEN SYSDATETIMEOFFSET() ELSE tgt.completed_at END,
    attempt_count = @attemptCount,
    duration_ms = @durationMs,
    error_code = @errorCode,
    error_message = @errorMessage,
    result_data = @resultData
WHEN NOT MATCHED THEN INSERT
    (job_id, job_type, status, source_path, source_folder, dedupe_key, trigger_source, started_at, attempt_count, duration_ms, error_code, error_message, result_data)
VALUES
    (@jobId, @jobType, @status, @sourcePath, @sourceFolder, @dedupeKey, @triggerSource, @startedAt, @attemptCount, @durationMs, @errorCode, @errorMessage, @resultData);
"@
        $params = @{
            jobId         = $JobId
            jobType       = $telemetryJobType
            status        = $Status
            sourcePath    = if ($SourcePath)    { $SourcePath }    else { $null }
            sourceFolder  = if ($SourceFolder)  { $SourceFolder }  else { $null }
            dedupeKey     = if ($DedupeKey)      { $DedupeKey }      else { $null }
            triggerSource = if ($TriggerSource) { $TriggerSource } else { $null }
            startedAt     = if ($startedAt) { $startedAt } else { $null }
            attemptCount  = $AttemptCount
            durationMs    = if ($null -ne $DurationMs) { $DurationMs } else { $null }
            errorCode     = if ($ErrorCode)     { $ErrorCode }     else { $null }
            errorMessage  = if ($ErrorMessage)  { $ErrorMessage }  else { $null }
            resultData    = if ($resultPayload) { $resultPayload } else { $null }
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'JOB_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='merge_processing_jobs'; jobId=$JobId; jobType=$telemetryJobType }
            return New-QCErrorResult -Code 'JOB_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='merge_processing_jobs'; jobId=$JobId; jobType=$telemetryJobType }
        }
        return New-QCSuccessResult -Code 'JOB_TELEMETRY_WRITTEN' -Message 'Job telemetry upserted.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected; jobType = $telemetryJobType }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'JOB_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='merge_processing_jobs'; jobId=$JobId; jobType=$telemetryJobType }
        return New-QCErrorResult -Code 'JOB_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='merge_processing_jobs'; jobId=$JobId; jobType=$telemetryJobType }
    }
}


function Test-QCPollRunTelemetryInsertShape {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Sql,
        [Parameter(Mandatory)][hashtable]$Parameters,
        [string[]]$RequiredParameterNames = @('durationMs','eventsFetched','eventsRelevant','candidatesCreated','jobsEnqueued','isReconciliation')
    )
    $mCols = [regex]::Match($Sql, 'INSERT\s+INTO\s+(?:dbo\.)?poll_runs\s*\(\s*(?<cols>[\s\S]*?)\s*\)\s*VALUES', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $mVals = [regex]::Match($Sql, 'VALUES\s*\(\s*(?<vals>[\s\S]*?)\s*\)\s*;?\s*$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $mCols.Success -or -not $mVals.Success) {
        return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_SQL_PARSE_FAILED' -Message 'Could not parse poll_runs insert statement.'
    }
    $cols = @($mCols.Groups['cols'].Value.Split(',') | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ })
    $paramRefs = [regex]::Matches($mVals.Groups['vals'].Value, '@([A-Za-z0-9_]+)') | ForEach-Object { $_.Groups[1].Value }
    $distinctCols = @($cols | Select-Object -Unique)
    if ($distinctCols.Count -ne $cols.Count) { return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_DUPLICATE_COLUMNS' -Message 'Duplicate columns detected in poll_runs insert.' -Data @{ columns = $cols } }
    if ($cols.Count -ne $paramRefs.Count) { return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_COLUMN_VALUE_MISMATCH' -Message "poll_runs insert has $($cols.Count) columns but $($paramRefs.Count) value parameters." -Data @{ columns = $cols.Count; params = $paramRefs.Count } }
    foreach ($rp in @($RequiredParameterNames)) { if (-not $Parameters.ContainsKey($rp)) { return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_REQUIRED_PARAM_MISSING' -Message "Missing required telemetry parameter: $rp" } }
    foreach ($pn in @($paramRefs)) { if (-not $Parameters.ContainsKey($pn)) { return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_PARAM_REF_MISSING' -Message "SQL references @${pn} but params has no key '${pn}'." } }
    return New-QCSuccessResult -Code 'POLL_RUN_TELEMETRY_INSERT_VALID' -Message 'poll_runs insert shape validated.' -Data @{ columnCount = $cols.Count; paramCount = $paramRefs.Count }
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
        [bool]$IsReconciliation = $false,
        [string]$WatcherName = 'qc_watcher',
        [string]$ServiceName = 'qc_pipeline',
        [Nullable[int]]$PassNumber,
        [string]$RunMode = 'audit',
        [string]$RunStatus = 'succeeded',
        [Nullable[decimal]]$TotalDurationSeconds,
        [Nullable[decimal]]$AuditQueryDurationSeconds,
        [Nullable[decimal]]$ReconciliationDurationSeconds,
        [Nullable[decimal]]$TriggerEvalDurationSeconds,
        [Nullable[decimal]]$DedupeDurationSeconds,
        [Nullable[decimal]]$QueueWriteDurationSeconds,
        [Nullable[decimal]]$DatabaseWriteDurationSeconds,
        [Nullable[decimal]]$CleanupDurationSeconds,
        [Nullable[decimal]]$SleepThrottleDurationSeconds,
        [int]$CandidateDocumentsEvaluated = 0,
        [int]$TriggerMatches = 0,
        [int]$JobsSkippedDedupe = 0,
        [int]$WarningCount = 0,
        [int]$ErrorCount = 0,
        [string]$ReconciliationReason,
        [string]$ReconciliationTriggerSource,
        [Nullable[int]]$DowntimeSeconds,
        [bool]$AuditGapDetected = $false,
        [string]$WatcherPhase,
        [Nullable[decimal]]$ThrottleWaitSeconds,
        [Nullable[int]]$QueueDepthSnapshot,
        [string]$RunId = $null,
        [string]$PassNumberSource = $null
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'POLL_RUN_TELEMETRY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
    try {
        $sql = @"
INSERT INTO poll_runs
    (started_at, completed_at, watermark_before, watermark_after, events_fetched, events_relevant, candidates_created, jobs_enqueued, duration_ms, error_message, is_reconciliation, watcher_name, service_name, pass_number, run_mode, run_status, total_duration_seconds, audit_query_duration_seconds, reconciliation_duration_seconds, trigger_eval_duration_seconds, dedupe_duration_seconds, queue_write_duration_seconds, database_write_duration_seconds, cleanup_duration_seconds, sleep_throttle_duration_seconds, candidate_documents_evaluated, trigger_matches, jobs_skipped_dedupe, warning_count, error_count, reconciliation_reason, reconciliation_trigger_source, downtime_seconds, audit_gap_detected, watcher_phase, throttle_wait_seconds, queue_depth_snapshot, pass_number_source)
VALUES
    (@startedAt, @completedAt, @watermarkBefore, @watermarkAfter, @eventsFetched, @eventsRelevant, @candidatesCreated, @jobsEnqueued, @durationMs, @errorMessage, @isReconciliation, @watcherName, @serviceName, @passNumber, @runMode, @runStatus, @totalDurationSeconds, @auditQueryDurationSeconds, @reconciliationDurationSeconds, @triggerEvalDurationSeconds, @dedupeDurationSeconds, @queueWriteDurationSeconds, @databaseWriteDurationSeconds, @cleanupDurationSeconds, @sleepThrottleDurationSeconds, @candidateDocumentsEvaluated, @triggerMatches, @jobsSkippedDedupe, @warningCount, @errorCount, @reconciliationReason, @reconciliationTriggerSource, @downtimeSeconds, @auditGapDetected, @watcherPhase, @throttleWaitSeconds, @queueDepthSnapshot, @passNumberSource)
"@
        $params = @{
            completedAt       = [DateTimeOffset]::Now
            eventsFetched     = $EventsFetched
            eventsRelevant    = $EventsRelevant
            candidatesCreated = $CandidatesCreated
            jobsEnqueued      = $JobsEnqueued
            durationMs        = if ($null -ne $DurationMs) { $DurationMs } else { 0 }
            errorMessage      = if ($ErrorMessage) { $ErrorMessage } else { $null }
            watermarkBefore   = if (-not [string]::IsNullOrWhiteSpace($WatermarkBefore)) { $WatermarkBefore.Trim() } else { $null }
            watermarkAfter    = if (-not [string]::IsNullOrWhiteSpace($WatermarkAfter)) { $WatermarkAfter.Trim() } else { $null }
            isReconciliation  = if ($IsReconciliation) { 1 } else { 0 }

            watcherName = if ($WatcherName) { $WatcherName } else { $null }
            serviceName = if ($ServiceName) { $ServiceName } else { $null }
            passNumber = if ($null -ne $PassNumber) { $PassNumber } else { $null }
            runMode = if ($RunMode) { $RunMode } else { $null }
            runStatus = if ($RunStatus) { $RunStatus } else { $null }
            totalDurationSeconds = if ($null -ne $TotalDurationSeconds) { $TotalDurationSeconds } else { $null }
            auditQueryDurationSeconds = if ($null -ne $AuditQueryDurationSeconds) { $AuditQueryDurationSeconds } else { $null }
            reconciliationDurationSeconds = if ($null -ne $ReconciliationDurationSeconds) { $ReconciliationDurationSeconds } else { $null }
            triggerEvalDurationSeconds = if ($null -ne $TriggerEvalDurationSeconds) { $TriggerEvalDurationSeconds } else { $null }
            dedupeDurationSeconds = if ($null -ne $DedupeDurationSeconds) { $DedupeDurationSeconds } else { $null }
            queueWriteDurationSeconds = if ($null -ne $QueueWriteDurationSeconds) { $QueueWriteDurationSeconds } else { $null }
            databaseWriteDurationSeconds = if ($null -ne $DatabaseWriteDurationSeconds) { $DatabaseWriteDurationSeconds } else { $null }
            cleanupDurationSeconds = if ($null -ne $CleanupDurationSeconds) { $CleanupDurationSeconds } else { $null }
            sleepThrottleDurationSeconds = if ($null -ne $SleepThrottleDurationSeconds) { $SleepThrottleDurationSeconds } else { $null }
            candidateDocumentsEvaluated = $CandidateDocumentsEvaluated
            triggerMatches = $TriggerMatches
            jobsSkippedDedupe = $JobsSkippedDedupe
            warningCount = $WarningCount
            errorCount = $ErrorCount
            reconciliationReason = if ($ReconciliationReason) { $ReconciliationReason } else { $null }
            reconciliationTriggerSource = if ($ReconciliationTriggerSource) { $ReconciliationTriggerSource } else { $null }
            downtimeSeconds = if ($null -ne $DowntimeSeconds) { $DowntimeSeconds } else { $null }
            auditGapDetected = if ($AuditGapDetected) { 1 } else { 0 }
            watcherPhase = if ($WatcherPhase) { $WatcherPhase } else { $null }
            throttleWaitSeconds = if ($null -ne $ThrottleWaitSeconds) { $ThrottleWaitSeconds } else { $null }
            queueDepthSnapshot = if ($null -ne $QueueDepthSnapshot) { $QueueDepthSnapshot } else { $null }
            passNumberSource = if ($PassNumberSource) { $PassNumberSource } else { $null }

        }
        try {
            $params['startedAt'] = ([DateTimeOffset]$params.completedAt).AddMilliseconds(-1 * [int]$params.durationMs)
        } catch {
            $params['startedAt'] = [DateTimeOffset]::Now.AddMilliseconds(-1 * [int]$params.durationMs)
        }
        $shape = Test-QCPollRunTelemetryInsertShape -Sql $sql -Parameters $params
        if (-not $shape.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'POLL_RUN_TELEMETRY_VALIDATION_FAILED' -Message $shape.Message -Data @{ operation='insert_poll_runs'; runId=$RunId; passNumber=$PassNumber; watcherName=$WatcherName; details=$shape.Data }
            return $shape
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'POLL_RUN_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='insert_poll_runs'; runId=$RunId; passNumber=$PassNumber; watcherName=$WatcherName }
            return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='insert_poll_runs'; runId=$RunId; passNumber=$PassNumber; watcherName=$WatcherName }
        }
        return New-QCSuccessResult -Code 'POLL_RUN_TELEMETRY_WRITTEN' -Message 'Poll run telemetry inserted.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'POLL_RUN_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='insert_poll_runs'; runId=$RunId; passNumber=$PassNumber; watcherName=$WatcherName }
        return New-QCErrorResult -Code 'POLL_RUN_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='insert_poll_runs'; runId=$RunId; passNumber=$PassNumber; watcherName=$WatcherName }
    }
}

function Write-QCDocumentStateHistoryRow {
    <#
    .SYNOPSIS
    Inserts one row into document_state_history (fire-and-forget).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$EventType,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$OldValue = '',
        [string]$NewValue = '',
        [string]$FieldName = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [Nullable[long]]$SourceAuditId = $null
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'STATE_HISTORY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'STATE_HISTORY_PLANNED' -Message 'Dry-run: state history not written.' -Data @{ written = $false; planned = $true }
    }
    $FolderPath = _QDB-NormalizeTelemetryPath -Path $FolderPath
    try {
        $sql = @"
INSERT INTO document_state_history
    (document_guid, document_name, folder_path, event_type, source_audit_id, old_value, new_value, field_name, changed_by_user, changed_by_username)
VALUES
    (@documentGuid, @documentName, @folderPath, @eventType, @sourceAuditId, @oldValue, @newValue, @fieldName, @changedByUser, @changedByUsername)
"@
        $params = @{
            documentGuid  = $DocumentGuid
            documentName  = if ($DocumentName) { $DocumentName } else { $null }
            folderPath    = if ($FolderPath) { $FolderPath } else { $null }
            eventType     = $EventType
            sourceAuditId = if ($null -ne $SourceAuditId) { $SourceAuditId } else { $null }
            oldValue      = if ($OldValue) { $OldValue } else { $null }
            newValue      = if ($NewValue) { $NewValue } else { $null }
            fieldName     = if ($FieldName) { $FieldName } else { $null }
            changedByUser = if ($null -ne $ChangedByUser) { $ChangedByUser } else { $null }
            changedByUsername = if ($ChangedByUsername) { [string]$ChangedByUsername } else { $null }
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            return New-QCErrorResult -Code 'STATE_HISTORY_WRITE_FAILED' -Message $dbRes.Message -Data @{ written = $false }
        }
        return New-QCSuccessResult -Code 'STATE_HISTORY_WRITTEN' -Message 'document_state_history row inserted.' -Data @{ written = $true }
    } catch {
        return New-QCErrorResult -Code 'STATE_HISTORY_EXCEPTION' -Message $_.Exception.Message -Data @{ written = $false }
    }
}

function Write-QCWorkflowEventRow {
    <#
    .SYNOPSIS
    Inserts one row into qc_workflow_events (audit, processor, and comment-sync state changes).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Nullable[long]]$RunId = $null,
        [string]$JobId = '',
        [string]$DocumentId = '',
        [Parameter(Mandatory)][string]$EventType,
        [string]$PreviousPwState = '',
        [string]$TargetPwState = '',
        [string]$DecisionCode = '',
        [string]$ProcessorVersion = '',
        [string]$QcReviewType = '',
        [string]$PayloadJson = '',
        [switch]$PlannedOnly
    )

    $writePlannedToDb = $false
    try {
        if ($Config.database -and $Config.database.logPlannedEventsInDryRun) {
            $writePlannedToDb = [bool]$Config.database.logPlannedEventsInDryRun
        }
    } catch { }

    $persistPlannedEvent = [bool]$PlannedOnly -and $writePlannedToDb
    if ($PlannedOnly -and -not $persistPlannedEvent) {
        return New-QCSuccessResult -Code 'QC_WORKFLOW_EVENT_PLANNED' -Message 'Workflow event not written (dry-run or DB disabled).' -Data @{
            planned = $true; eventType = $EventType; decisionCode = $DecisionCode
        }
    }
    if (-not $PlannedOnly -and -not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_WORKFLOW_EVENT_PLANNED' -Message 'Workflow event not written (dry-run or DB disabled).' -Data @{
            planned = $true; eventType = $EventType; decisionCode = $DecisionCode
        }
    }

    $eventTypeValue = [string]$EventType
    if ($persistPlannedEvent) { $eventTypeValue = ($eventTypeValue + '_PLANNED') }

    try {
        $sql = @"
INSERT INTO qc_workflow_events
    (run_id, job_id, document_id, event_type, previous_pw_state, target_pw_state, decision_code, processor_version, qc_review_type, payload_json)
VALUES
    (@runId, @jobId, @documentId, @eventType, @prev, @target, @decisionCode, @procVer, @qcReviewType, @payload)
"@
        $params = @{
            runId = if ($null -ne $RunId -and $RunId -gt 0) { $RunId } else { [DBNull]::Value }
            jobId = if ($JobId) { $JobId } else { [DBNull]::Value }
            documentId = if ($DocumentId) { $DocumentId } else { [DBNull]::Value }
            eventType = $eventTypeValue
            prev = if ($PreviousPwState) { $PreviousPwState } else { [DBNull]::Value }
            target = if ($TargetPwState) { $TargetPwState } else { [DBNull]::Value }
            decisionCode = if ($DecisionCode) { $DecisionCode } else { [DBNull]::Value }
            procVer = if ($ProcessorVersion) { $ProcessorVersion } else { [DBNull]::Value }
            qcReviewType = if ($QcReviewType) { $QcReviewType } else { [DBNull]::Value }
            payload = if ($PayloadJson) { $PayloadJson } else { [DBNull]::Value }
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            return New-QCErrorResult -Code 'QC_WORKFLOW_EVENT_FAILED' -Message $dbRes.Message -Data @{ written = $false; eventType = $eventTypeValue }
        }
        return New-QCSuccessResult -Code 'QC_WORKFLOW_EVENT_WRITTEN' -Message 'Workflow event inserted.' -Data @{ written = $true; eventType = $eventTypeValue }
    } catch {
        return New-QCErrorResult -Code 'QC_WORKFLOW_EVENT_EXCEPTION' -Message $_.Exception.Message -Data @{ written = $false; eventType = $eventTypeValue }
    }
}

function Ensure-QCTransitionEvent {
    <#
    .SYNOPSIS
    Returns an existing transition_events row for the same state change when possible; inserts only for a new cycle.
    .DESCRIPTION
    Reuses a pending (notification_sent=0) row for echo audits. Reuses a sent row only when the same
    trigger_audit_id is seen again. After notification_sent=1 with a different audit/job, inserts a new row
    so a later QC cycle to the same target state can notify again.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$TransitionType,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$FromValue = '',
        [string]$ToValue = '',
        [string]$JobId = '',
        [string]$JobType = '',
        [Nullable[long]]$TriggerAuditId = $null,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'TRANSITION_EVENT_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false; transitionId = $null; reused = $false }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'TRANSITION_EVENT_PLANNED' -Message 'Dry-run: transition event not written.' -Data @{ written = $false; planned = $true; transitionId = $null; reused = $false }
    }

    $fromNorm = if ($FromValue) { ([string]$FromValue).Trim() } else { '' }
    $toNorm = if ($ToValue) { ([string]$ToValue).Trim() } else { '' }
    $typeNorm = ([string]$TransitionType).Trim()

    try {
        if ($null -ne $TriggerAuditId -and $TriggerAuditId -gt 0) {
            $auditRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 id, notification_sent
FROM transition_events
WHERE document_guid = @documentGuid
  AND transition_type = @transitionType
  AND ISNULL(from_value, '') = @fromValue
  AND ISNULL(to_value, '') = @toValue
  AND trigger_audit_id = @triggerAuditId
ORDER BY id DESC
"@ -Parameters @{
                documentGuid = $DocumentGuid
                transitionType = $typeNorm
                fromValue = $fromNorm
                toValue = $toNorm
                triggerAuditId = [long]$TriggerAuditId
            }
            if ($auditRes.IsSuccess -and $auditRes.Data.table -and $auditRes.Data.table.Rows.Count -gt 0) {
                $r = $auditRes.Data.table.Rows[0]
                $existingId = [int]$r.id
                return New-QCSuccessResult -Code 'TRANSITION_EVENT_REUSED' -Message 'Existing transition_events row reused (same audit).' -Data @{
                    written = $false; reused = $true; transitionId = $existingId
                    notificationSent = [bool]$r.notification_sent
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($JobId)) {
            $jobRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 id, notification_sent
FROM transition_events
WHERE document_guid = @documentGuid
  AND transition_type = @transitionType
  AND ISNULL(from_value, '') = @fromValue
  AND ISNULL(to_value, '') = @toValue
  AND job_id = @jobId
ORDER BY id DESC
"@ -Parameters @{
                documentGuid = $DocumentGuid
                transitionType = $typeNorm
                fromValue = $fromNorm
                toValue = $toNorm
                jobId = [string]$JobId
            }
            if ($jobRes.IsSuccess -and $jobRes.Data.table -and $jobRes.Data.table.Rows.Count -gt 0) {
                $r = $jobRes.Data.table.Rows[0]
                $existingId = [int]$r.id
                return New-QCSuccessResult -Code 'TRANSITION_EVENT_REUSED' -Message 'Existing transition_events row reused (same job).' -Data @{
                    written = $false; reused = $true; transitionId = $existingId
                    notificationSent = [bool]$r.notification_sent
                }
            }
        }

        $latestRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 id, notification_sent, trigger_audit_id
FROM transition_events
WHERE document_guid = @documentGuid
  AND transition_type = @transitionType
  AND ISNULL(from_value, '') = @fromValue
  AND ISNULL(to_value, '') = @toValue
ORDER BY id DESC
"@ -Parameters @{
            documentGuid = $DocumentGuid
            transitionType = $typeNorm
            fromValue = $fromNorm
            toValue = $toNorm
        }
        if ($latestRes.IsSuccess -and $latestRes.Data.table -and $latestRes.Data.table.Rows.Count -gt 0) {
            $r = $latestRes.Data.table.Rows[0]
            $sent = $false
            try { $sent = [bool]$r.notification_sent } catch { }
            if (-not $sent) {
                $existingId = [int]$r.id
                return New-QCSuccessResult -Code 'TRANSITION_EVENT_REUSED' -Message 'Existing pending transition_events row reused.' -Data @{
                    written = $false; reused = $true; transitionId = $existingId; notificationSent = $false
                }
            }
        }
    } catch { }

    return Write-QCTransitionEvent -Config $Config -DocumentGuid $DocumentGuid -TransitionType $TransitionType `
        -DocumentName $DocumentName -FolderPath $FolderPath -FromValue $FromValue -ToValue $ToValue `
        -JobId $JobId -JobType $JobType -TriggerAuditId $TriggerAuditId -ChangedByUser $ChangedByUser `
        -ChangedByUsername $ChangedByUsername
}

function Test-QCTransitionEventNotificationSent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][int]$TransitionId
    )

    if ($TransitionId -le 0) { return $false }
    if (-not (_QDB-IsEnabled -Config $Config)) { return $false }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 notification_sent FROM transition_events WHERE id = @id
"@ -Parameters @{ id = $TransitionId }
        if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
            $r = $res.Data.table.Rows[0]
            if (-not ($r.notification_sent -is [DBNull])) { return [bool]$r.notification_sent }
        }
    } catch { }
    return $false
}

function Write-QCTransitionEvent {
    <#
    .SYNOPSIS
    Inserts a business-level transition_events row and returns the new id when possible.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$TransitionType,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$FromValue = '',
        [string]$ToValue = '',
        [string]$JobId = '',
        [string]$JobType = '',
        [Nullable[long]]$TriggerAuditId = $null,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'TRANSITION_EVENT_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false; transitionId = $null }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'TRANSITION_EVENT_PLANNED' -Message 'Dry-run: transition event not written.' -Data @{ written = $false; planned = $true; transitionId = $null }
    }
    $FolderPath = _QDB-NormalizeTelemetryPath -Path $FolderPath
    try {
        $sql = @"
INSERT INTO transition_events
    (document_guid, document_name, folder_path, transition_type, from_value, to_value, trigger_audit_id, job_id, job_type, changed_by_user, changed_by_username)
OUTPUT INSERTED.id
VALUES
    (@documentGuid, @documentName, @folderPath, @transitionType, @fromValue, @toValue, @triggerAuditId, @jobId, @jobType, @changedByUser, @changedByUsername)
"@
        $params = @{
            documentGuid   = $DocumentGuid
            documentName   = if ($DocumentName) { $DocumentName } else { $null }
            folderPath     = if ($FolderPath) { $FolderPath } else { $null }
            transitionType = $TransitionType
            fromValue      = if ($FromValue) { $FromValue } else { $null }
            toValue        = if ($ToValue) { $ToValue } else { $null }
            triggerAuditId = if ($null -ne $TriggerAuditId) { $TriggerAuditId } else { $null }
            jobId          = if ($JobId) { $JobId } else { $null }
            jobType        = if ($JobType) { $JobType } else { $null }
            changedByUser  = if ($null -ne $ChangedByUser) { $ChangedByUser } else { $null }
            changedByUsername = if ($ChangedByUsername) { [string]$ChangedByUsername } else { $null }
        }
        $dbRes = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
        $transitionId = $null
        if ($dbRes.IsSuccess -and $null -ne $dbRes.Data.value) {
            try { $transitionId = [int]$dbRes.Data.value } catch { $transitionId = $null }
        }
        if (-not $dbRes.IsSuccess) {
            return New-QCErrorResult -Code 'TRANSITION_EVENT_WRITE_FAILED' -Message $dbRes.Message -Data @{ written = $false; transitionId = $null }
        }
        return New-QCSuccessResult -Code 'TRANSITION_EVENT_WRITTEN' -Message 'transition_events row inserted.' -Data @{ written = $true; transitionId = $transitionId }
    } catch {
        return New-QCErrorResult -Code 'TRANSITION_EVENT_EXCEPTION' -Message $_.Exception.Message -Data @{ written = $false; transitionId = $null }
    }
}

function Update-QCTransitionEventNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][int]$TransitionId,
        [bool]$NotificationSent = $true,
        [string]$NotificationId = ''
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) { return }
    try {
        $sql = @"
UPDATE transition_events
SET notification_sent = @sent, notification_id = @notificationId
WHERE id = @id
"@
        $params = @{
            id = $TransitionId
            sent = if ($NotificationSent) { 1 } else { 0 }
            notificationId = if ($NotificationId) { $NotificationId } else { $null }
        }
        Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params | Out-Null
    } catch { }
}

function Get-QCTransitionEventActor {
    <#
    .SYNOPSIS
    Returns changed_by_user and changed_by_username for a transition_events row.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][int]$TransitionId
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return @{ changedByUser = $null; changedByUsername = '' }
    }
    try {
        $sql = @"
SELECT changed_by_user, changed_by_username
FROM transition_events
WHERE id = @id
"@
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters @{ id = $TransitionId }
        if (-not $res.IsSuccess -or -not $res.Data.rows -or $res.Data.rows.Count -lt 1) {
            return @{ changedByUser = $null; changedByUsername = '' }
        }
        $row = $res.Data.rows[0]
        $user = $null
        if ($null -ne $row.changed_by_user -and -not ($row.changed_by_user -is [DBNull])) {
            try { $user = [int]$row.changed_by_user } catch { $user = $null }
        }
        $username = ''
        if ($null -ne $row.changed_by_username -and -not ($row.changed_by_username -is [DBNull])) {
            $username = [string]$row.changed_by_username
        }
        return @{ changedByUser = $user; changedByUsername = $username }
    } catch {
        return @{ changedByUser = $null; changedByUsername = '' }
    }
}

function Get-QCAuditEventActor {
    <#
    .SYNOPSIS
    Returns pw_userno and username for an audit_events row (audit trigger actor source).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][long]$AuditEventId
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return @{ changedByUser = $null; changedByUsername = '' }
    }
    try {
        $sql = @"
SELECT ae.pw_userno, pu.pw_username
FROM audit_events ae
LEFT JOIN pw_users pu ON pu.pw_userno = ae.pw_userno
WHERE ae.id = @id
"@
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters @{ id = $AuditEventId }
        if (-not $res.IsSuccess -or -not $res.Data.rows -or $res.Data.rows.Count -lt 1) {
            return @{ changedByUser = $null; changedByUsername = '' }
        }
        $row = $res.Data.rows[0]
        $user = $null
        if ($null -ne $row.pw_userno -and -not ($row.pw_userno -is [DBNull])) {
            try { $user = [int]$row.pw_userno } catch { $user = $null }
        }
        $username = ''
        if ($null -ne $row.pw_username -and -not ($row.pw_username -is [DBNull])) {
            $username = [string]$row.pw_username
        }
        return @{ changedByUser = $user; changedByUsername = $username }
    } catch {
        return @{ changedByUser = $null; changedByUsername = '' }
    }
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
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'NOTIFICATION_TELEMETRY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
    $FolderPath = _QDB-NormalizeTelemetryPath -Path $FolderPath
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
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'NOTIFICATION_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='insert_notification_log'; eventType=$EventType }
            return New-QCErrorResult -Code 'NOTIFICATION_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='insert_notification_log'; eventType=$EventType }
        }
        return New-QCSuccessResult -Code 'NOTIFICATION_TELEMETRY_WRITTEN' -Message 'Notification telemetry inserted.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'NOTIFICATION_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='insert_notification_log'; eventType=$EventType }
        return New-QCErrorResult -Code 'NOTIFICATION_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='insert_notification_log'; eventType=$EventType }
    }
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
        [string]$CheckerEmail,
        [string]$QcReviewType,
        [string]$QcAssignedTo,
        [string]$PwStateName,
        [string]$QcStage,
        [string]$QcStatus,
        [string]$LastAuditEventAt,
        [string]$FileModifiedAt,
        [switch]$SetOwnershipFromProjectWise
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
    if ([string]::IsNullOrWhiteSpace($DocumentGuid) -or $DocumentGuid.Trim().Length -lt 5) {
        return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'document_guid is missing or too short for a ProjectWise GUID.' -Data @{ written = $false; documentGuid = $DocumentGuid }
    }
    $FolderPath = _QDB-NormalizeTelemetryPath -Path $FolderPath
    if (-not (Test-QCSheetIndexFolderPath -FolderPath $FolderPath)) {
        return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'folder_path is outside CADD/Sheets; sheet_index tracks Sheets folders only.' -Data @{ written = $false; folderPath = $FolderPath }
    }
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
    checker_email = CASE WHEN @setOwnership = 1 THEN COALESCE(@checkerEmail, tgt.checker_email) ELSE COALESCE(@checkerEmail, tgt.checker_email) END,
    qc_review_type = CASE WHEN @setOwnership = 1 THEN COALESCE(@qcReviewType, tgt.qc_review_type) ELSE COALESCE(@qcReviewType, tgt.qc_review_type) END,
    qc_assigned_to = CASE WHEN @setOwnership = 1 THEN COALESCE(@qcAssignedTo, tgt.qc_assigned_to) ELSE COALESCE(@qcAssignedTo, tgt.qc_assigned_to) END,
    pw_state_name = CASE WHEN @setOwnership = 1 THEN @pwStateName ELSE COALESCE(@pwStateName, tgt.pw_state_name) END,
    qc_stage = COALESCE(@qcStage, tgt.qc_stage),
    qc_status = CASE WHEN @setOwnership = 1 THEN @qcStatus ELSE COALESCE(@qcStatus, tgt.qc_status) END,
    last_updated_at = SYSDATETIMEOFFSET(),
    last_audit_event_at = COALESCE(@lastAuditEventAt, tgt.last_audit_event_at),
    file_modified_at = COALESCE(@fileModifiedAt, tgt.file_modified_at)
WHEN NOT MATCHED THEN INSERT
    (document_guid, document_name, document_number, folder_path, project_name, watch_root,
     extension, source_type, designer_email, reviewer_email, checker_email, qc_review_type, qc_assigned_to,
     pw_state_name, qc_stage, qc_status, last_audit_event_at, file_modified_at)
VALUES
    (@docGuid, @docName, @docNumber, @folderPath, @projectName, @watchRoot,
     @extension, @sourceType, @designerEmail, @reviewerEmail, @checkerEmail, @qcReviewType, @qcAssignedTo,
     @pwStateName, @qcStage, @qcStatus, @lastAuditEventAt, @fileModifiedAt);
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
            checkerEmail     = if ($CheckerEmail)      { $CheckerEmail }      else { $null }
            qcReviewType     = if ($QcReviewType)      { $QcReviewType }      else { $null }
            qcAssignedTo     = if ($QcAssignedTo)      { $QcAssignedTo }      else { $null }
            pwStateName      = if ($PwStateName)       { $PwStateName }       else { $null }
            qcStage          = if ($QcStage)           { $QcStage }           else { $null }
            qcStatus         = if ($QcStatus)          { $QcStatus }          else { $null }
            lastAuditEventAt = if ($LastAuditEventAt)  { $LastAuditEventAt }  else { $null }
            fileModifiedAt   = if ($FileModifiedAt)    { $FileModifiedAt }    else { $null }
            setOwnership     = $setOwnership
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='merge_sheet_index'; documentGuid=$DocumentGuid }
            return New-QCErrorResult -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='merge_sheet_index'; documentGuid=$DocumentGuid }
        }
        return New-QCSuccessResult -Code 'SHEET_INDEX_WRITTEN' -Message 'Sheet index upserted.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='merge_sheet_index'; documentGuid=$DocumentGuid }
        return New-QCErrorResult -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='merge_sheet_index'; documentGuid=$DocumentGuid }
    }
}

function Update-QCSheetIndexPwStateName {
    <#
    .SYNOPSIS
    Updates only pw_state_name for one sheet_index row. Does not touch email columns.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$PwStateName
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
    try {
        $sql = @"
UPDATE sheet_index
SET pw_state_name = @pwStateName,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@
        $params = @{
            docGuid     = $DocumentGuid
            pwStateName = $PwStateName
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='update_sheet_index_pw_state'; documentGuid=$DocumentGuid }
            return New-QCErrorResult -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='update_sheet_index_pw_state'; documentGuid=$DocumentGuid }
        }
        return New-QCSuccessResult -Code 'SHEET_INDEX_WRITTEN' -Message 'Sheet index state updated.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='update_sheet_index_pw_state'; documentGuid=$DocumentGuid }
        return New-QCErrorResult -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='update_sheet_index_pw_state'; documentGuid=$DocumentGuid }
    }
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
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
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
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='update_sheet_index_qc_pdf'; sourceDocumentGuid=$SourceDocumentGuid }
            return New-QCErrorResult -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='update_sheet_index_qc_pdf'; sourceDocumentGuid=$SourceDocumentGuid }
        }
        return New-QCSuccessResult -Code 'SHEET_INDEX_WRITTEN' -Message 'Sheet index QC PDF link updated.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='update_sheet_index_qc_pdf'; sourceDocumentGuid=$SourceDocumentGuid }
        return New-QCErrorResult -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='update_sheet_index_qc_pdf'; sourceDocumentGuid=$SourceDocumentGuid }
    }
}

function Write-QCSheetIndexBatch {
    <#
    .SYNOPSIS
    Batch upsert into sheet_index for reconciliation scans.
    .DESCRIPTION
    Uses a staging table and one MERGE to reduce round trips.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable[]]$Rows
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'POLL_RUN_TELEMETRY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $Rows = @($Rows | ForEach-Object {
        $row = $_
        $fp = if ($row.folderPath) { _QDB-NormalizeTelemetryPath -Path ([string]$row.folderPath) } else { $null }
        if ($fp) { $row.folderPath = $fp }
        $row
    } | Where-Object {
        $g = if ($_.documentGuid) { [string]$_.documentGuid } else { '' }
        $fp = if ($_.folderPath) { [string]$_.folderPath } else { '' }
        -not [string]::IsNullOrWhiteSpace($g) -and $g.Trim().Length -ge 5 -and (Test-QCSheetIndexFolderPath -FolderPath $fp)
    })
    if ($Rows.Count -eq 0) { return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'No rows with valid document_guid length.' -Data @{ written = $false } }
    try {
        $sessRes = New-QCDatabaseSession -Config $Config
        if (-not $sessRes.IsSuccess) { return }
        $sess = $sessRes.Data.session
        $conn = $sess.connection
        try {
            $createSql = @"
IF OBJECT_ID('tempdb..#sheet_index_stage') IS NOT NULL DROP TABLE #sheet_index_stage;
CREATE TABLE #sheet_index_stage (
    document_guid NVARCHAR(40) NOT NULL,
    document_name NVARCHAR(500) NOT NULL,
    folder_path   NVARCHAR(1000) NOT NULL,
    watch_root    NVARCHAR(500) NULL,
    source_type   NVARCHAR(10) NULL,
    designer_email NVARCHAR(200) NULL,
    reviewer_email NVARCHAR(200) NULL,
    checker_email NVARCHAR(200) NULL,
    qc_review_type NVARCHAR(100) NULL,
    qc_assigned_to NVARCHAR(200) NULL,
    qc_status NVARCHAR(50) NULL,
    pw_state_name  NVARCHAR(100) NULL
);
"@
            [void](Invoke-QCDatabaseNonQueryWithConnection -Connection $conn -Sql $createSql -Parameters @{} -CommandTimeout 120)

            $sheetParamsPerRow = 12
            $chunkSize = _QDB-GetMaxRowsForSqlParameters -ParametersPerRow $sheetParamsPerRow
            if ($chunkSize -gt 200) { $chunkSize = 200 }
            for ($i = 0; $i -lt $Rows.Count; $i += $chunkSize) {
                $chunk = @($Rows[$i..[Math]::Min($i + $chunkSize - 1, $Rows.Count - 1)])
                $sb = New-Object System.Text.StringBuilder
                $params = @{}
                for ($r = 0; $r -lt $chunk.Count; $r++) {
                    $row = $chunk[$r]
                    if ($r -gt 0) { [void]$sb.AppendLine(',') }
                    [void]$sb.Append(("(@docGuid{0},@docName{0},@folderPath{0},@watchRoot{0},@sourceType{0},@designerEmail{0},@reviewerEmail{0},@checkerEmail{0},@qcReviewType{0},@qcAssignedTo{0},@qcStatus{0},@pwStateName{0})" -f $r))
                    $params["docGuid$r"] = [string]$row.documentGuid
                    $params["docName$r"] = [string]$row.documentName
                    $params["folderPath$r"] = [string]$row.folderPath
                    $params["watchRoot$r"] = if ($row.watchRoot) { [string]$row.watchRoot } else { $null }
                    $params["sourceType$r"] = if ($row.sourceType) { [string]$row.sourceType } else { $null }
                    $params["designerEmail$r"] = if ($row.designerEmail) { [string]$row.designerEmail } else { $null }
                    $params["reviewerEmail$r"] = if ($row.reviewerEmail) { [string]$row.reviewerEmail } else { $null }
                    $params["checkerEmail$r"] = if ($row.checkerEmail) { [string]$row.checkerEmail } else { $null }
                    $params["qcReviewType$r"] = if ($row.qcReviewType) { [string]$row.qcReviewType } else { $null }
                    $params["qcAssignedTo$r"] = if ($row.qcAssignedTo) { [string]$row.qcAssignedTo } else { $null }
                    $params["qcStatus$r"] = if ($row.qcStatus) { [string]$row.qcStatus } else { $null }
                    $params["pwStateName$r"] = if ($row.pwStateName) { [string]$row.pwStateName } else { $null }
                }
                $insSql = @"
INSERT INTO #sheet_index_stage
    (document_guid, document_name, folder_path, watch_root, source_type, designer_email, reviewer_email, checker_email, qc_review_type, qc_assigned_to, qc_status, pw_state_name)
VALUES
$($sb.ToString());
"@
                [void](Invoke-QCDatabaseNonQueryWithConnection -Connection $conn -Sql $insSql -Parameters $params -CommandTimeout 120)
            }

            $mergeSql = @"
MERGE sheet_index AS tgt
USING #sheet_index_stage AS src
ON tgt.document_guid = src.document_guid
WHEN MATCHED THEN UPDATE SET
    document_name = src.document_name,
    folder_path = src.folder_path,
    watch_root = COALESCE(src.watch_root, tgt.watch_root),
    source_type = COALESCE(src.source_type, tgt.source_type),
    designer_email = COALESCE(src.designer_email, tgt.designer_email),
    reviewer_email = COALESCE(src.reviewer_email, tgt.reviewer_email),
    checker_email = COALESCE(src.checker_email, tgt.checker_email),
    qc_review_type = COALESCE(src.qc_review_type, tgt.qc_review_type),
    qc_assigned_to = COALESCE(src.qc_assigned_to, tgt.qc_assigned_to),
    qc_status = COALESCE(src.qc_status, tgt.qc_status),
    pw_state_name = COALESCE(src.pw_state_name, tgt.pw_state_name),
    last_updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT
    (document_guid, document_name, folder_path, watch_root, source_type, designer_email, reviewer_email, checker_email, qc_review_type, qc_assigned_to, qc_status, pw_state_name)
VALUES
    (src.document_guid, src.document_name, src.folder_path, src.watch_root, src.source_type, src.designer_email, src.reviewer_email, src.checker_email, src.qc_review_type, src.qc_assigned_to, src.qc_status, src.pw_state_name);
"@
            [void](Invoke-QCDatabaseNonQueryWithConnection -Connection $conn -Sql $mergeSql -Parameters @{} -CommandTimeout 120)
        } finally {
            try { $sess.Dispose.Invoke() } catch { }
        }
    } catch { }
}

function Get-QCDocumentFolderCache {
    <#
    .SYNOPSIS
    Returns document_guid -> folder_path from document_activity for audit folder resolution.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $cache = @{}
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'DOC_FOLDER_CACHE_SKIPPED' -Message 'Database disabled.' -Data @{ cache = $cache; count = 0 }
    }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT document_guid, folder_path
FROM document_activity
WHERE NULLIF(LTRIM(RTRIM(document_guid)), '') IS NOT NULL
  AND NULLIF(LTRIM(RTRIM(folder_path)), '') IS NOT NULL
"@
        if ($res.IsSuccess -and $res.Data -and $res.Data.table) {
            foreach ($row in @(_QDB-ConvertDataTableToRowHashtables -Table $res.Data.table)) {
                $g = [string]$row.document_guid
                $fp = [string]$row.folder_path
                if ($g -and $fp) { $cache[$g.Trim().ToLowerInvariant()] = $fp }
            }
        }
    } catch { }
    return New-QCSuccessResult -Code 'DOC_FOLDER_CACHE_OK' -Message "Loaded $($cache.Count) cached document folders." -Data @{ cache = $cache; count = $cache.Count }
}

function Get-QCUnprocessedAuditEvents {
    <#
    .SYNOPSIS
    Loads audit_events rows not yet processed by the watcher trigger pipeline.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [int]$MaxRows = 500,
        [int[]]$ActionCodes = @(1001, 1002, 1003, 1006, 1007, 1012, 1015, 1020),
        [long[]]$ExcludeEventIds = @()
    )

    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'AUDIT_UNPROCESSED_SKIPPED' -Message 'Database disabled.' -Data @{ rows = @(); count = 0 }
    }
    if ($MaxRows -lt 1) { $MaxRows = 500 }
    $codes = @($ActionCodes | Where-Object { $_ -gt 0 } | Select-Object -Unique)
    if ($codes.Count -eq 0) {
        return New-QCSuccessResult -Code 'AUDIT_UNPROCESSED_NONE' -Message 'No action codes configured.' -Data @{ rows = @(); count = 0 }
    }
    $codeList = ($codes | ForEach-Object { [string][int]$_ }) -join ','
    $excludeIds = @($ExcludeEventIds | Where-Object { $_ -gt 0 } | ForEach-Object { [string][long]$_ } | Select-Object -Unique)
    $excludeSql = ''
    if ($excludeIds.Count -gt 0) { $excludeSql = "`n  AND ae.id NOT IN ($($excludeIds -join ','))" }
    $sql = @"
SELECT TOP ($MaxRows)
    ae.id, ae.pw_acttime, ae.pw_action, ae.pw_action_name, ae.pw_objtype, ae.pw_objno, ae.pw_objguid, ae.pw_parentguid,
    ae.pw_userno, pu.pw_username, ae.pw_itemname, ae.pw_itemdesc, ae.pw_textparam, ae.resolved_folder, ae.candidate_type
FROM audit_events ae
LEFT JOIN pw_users pu ON pu.pw_userno = ae.pw_userno
WHERE ae.processed = 0
  AND ae.pw_action IN ($codeList)$excludeSql
ORDER BY CASE WHEN ae.pw_action = 1012 THEN 0 ELSE 1 END, ae.pw_acttime ASC, ae.pw_objguid ASC, ae.id ASC
"@
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql
        if (-not $res.IsSuccess) {
            return New-QCFailureResult -Code 'AUDIT_UNPROCESSED_QUERY_FAILED' -Message $res.Message -Data @{ rows = @(); count = 0 }
        }
        $rows = @()
        if ($res.Data -and $res.Data.table) {
            $rows = @(_QDB-ConvertDataTableToRowHashtables -Table $res.Data.table)
        }
        return New-QCSuccessResult -Code 'AUDIT_UNPROCESSED_OK' -Message "Loaded $($rows.Count) unprocessed audit events." -Data @{ rows = $rows; count = $rows.Count }
    } catch {
        return New-QCFailureResult -Code 'AUDIT_UNPROCESSED_EXCEPTION' -Message $_.Exception.Message -Data @{ rows = @(); count = 0 }
    }
}

function Update-QCAuditEventsResolvedFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Updates
    )

    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return New-QCSuccessResult -Code 'AUDIT_FOLDER_UPDATE_SKIPPED' -Message 'Database disabled.' -Data @{ updated = 0 } }
    if (-not $Updates -or $Updates.Count -eq 0) { return New-QCSuccessResult -Code 'AUDIT_FOLDER_UPDATE_NONE' -Message 'No updates.' -Data @{ updated = 0 } }

    $updated = 0
    foreach ($u in @($Updates)) {
        $id = 0
        try { $id = [long]$u.id } catch { continue }
        if ($id -le 0) { continue }
        $folder = [string]$u.resolvedFolder
        $ctype = if ($u.candidateType) { [string]$u.candidateType } else { $null }
        if ([string]::IsNullOrWhiteSpace($folder)) { continue }
        $folder = _QDB-NormalizeTelemetryPath -Path $folder
        if ([string]::IsNullOrWhiteSpace($folder)) { continue }
        try {
            $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE audit_events
SET resolved_folder = @folder, candidate_type = COALESCE(@ctype, candidate_type)
WHERE id = @id AND processed = 0
"@ -Parameters @{ id = $id; folder = $folder; ctype = $ctype }
            if ($res.IsSuccess) { $updated += [int]$res.Data.rowsAffected }
        } catch { }
    }
    return New-QCSuccessResult -Code 'AUDIT_FOLDER_UPDATE_OK' -Message "Updated $updated audit event folder paths." -Data @{ updated = $updated }
}

function Mark-QCAuditEventsProcessed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][long[]]$EventIds
    )

    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return New-QCSuccessResult -Code 'AUDIT_MARK_PROCESSED_SKIPPED' -Message 'Database disabled.' -Data @{ marked = 0 } }
    $ids = @($EventIds | Where-Object { $_ -gt 0 } | Select-Object -Unique)
    if ($ids.Count -eq 0) { return New-QCSuccessResult -Code 'AUDIT_MARK_PROCESSED_NONE' -Message 'No ids.' -Data @{ marked = 0 } }

    $marked = 0
    $chunkSize = 200
    for ($i = 0; $i -lt $ids.Count; $i += $chunkSize) {
        $chunk = @($ids[$i..[Math]::Min($i + $chunkSize - 1, $ids.Count - 1)])
        $paramNames = @()
        $params = @{}
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $paramNames += "@id$j"
            $params["id$j"] = $chunk[$j]
        }
        $inList = $paramNames -join ','
        try {
            $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql "UPDATE audit_events SET processed = 1 WHERE id IN ($inList) AND processed = 0" -Parameters $params
            if ($res.IsSuccess) { $marked += [int]$res.Data.rowsAffected }
        } catch { }
    }
    return New-QCSuccessResult -Code 'AUDIT_MARK_PROCESSED_OK' -Message "Marked $marked audit events processed." -Data @{ marked = $marked }
}

function Upsert-QCDocumentActivityFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [string]$DocumentName = '',
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$LastAction = '',
        [int]$LastActionCode = 0,
        [string]$LastActionTime = ''
    )

    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }
    $g = ([string]$DocumentGuid).Trim()
    if ([string]::IsNullOrWhiteSpace($g)) { return }
    $fp = _QDB-NormalizeTelemetryPath -Path $FolderPath
    if ([string]::IsNullOrWhiteSpace($fp)) { return }
    try {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
MERGE document_activity AS tgt
USING (SELECT @guid AS document_guid) AS src
ON tgt.document_guid = src.document_guid
WHEN MATCHED THEN UPDATE SET
    document_name = COALESCE(NULLIF(LTRIM(RTRIM(@name)), ''), tgt.document_name),
    folder_path = @folder,
    last_action = COALESCE(NULLIF(LTRIM(RTRIM(@action)), ''), tgt.last_action),
    last_action_code = CASE WHEN @actionCode > 0 THEN @actionCode ELSE tgt.last_action_code END,
    last_action_time = COALESCE(NULLIF(LTRIM(RTRIM(@actionTime)), ''), tgt.last_action_time),
    last_seen = SYSDATETIMEOFFSET(),
    total_events = tgt.total_events + 1
WHEN NOT MATCHED THEN INSERT
    (document_guid, document_name, folder_path, last_action, last_action_code, last_action_time, total_events, first_seen, last_seen)
VALUES
    (@guid, @name, @folder, @action, @actionCode, @actionTime, 1, SYSDATETIMEOFFSET(), SYSDATETIMEOFFSET());
"@ -Parameters @{
            guid = $g; name = $DocumentName; folder = $fp
            action = $LastAction; actionCode = $LastActionCode; actionTime = $LastActionTime
        })
    } catch { }
}

function Get-QCWatcherStateValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$StateKey
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }
    try {
        $res = Invoke-QCDatabaseScalar -Config $Config -Sql 'SELECT state_value FROM watcher_state WHERE state_key = @k' -Parameters @{ k = $StateKey }
        if ($res.IsSuccess -and $res.Data.value) { return [string]$res.Data.value }
    } catch { }
    return $null
}

function Set-QCWatcherStateValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$StateKey,
        [Parameter(Mandatory)][string]$StateValue
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $false }
    try {
        $res = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
MERGE watcher_state AS tgt
USING (SELECT @k AS state_key, @v AS state_value) AS src
ON tgt.state_key = src.state_key
WHEN MATCHED THEN UPDATE SET state_value = src.state_value, updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (state_key, state_value) VALUES (src.state_key, src.state_value);
"@ -Parameters @{ k = $StateKey; v = $StateValue }
        return [bool]$res.IsSuccess
    } catch { return $false }
}

function Get-QCAuditWatermarkUtc {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $raw = Get-QCWatcherStateValue -Config $Config -StateKey 'audit_watermark_utc'
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    try {
        $s = $raw.Trim().TrimEnd('Z')
        return [DateTime]::ParseExact(
            $s,
            'yyyy-MM-dd HH:mm:ss',
            $null,
            [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        )
    } catch {
        try { return [DateTime]::Parse($raw).ToUniversalTime() } catch { return $null }
    }
}

function Set-QCAuditWatermarkUtc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][DateTime]$WatermarkUtc
    )
    $value = $WatermarkUtc.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
    return (Set-QCWatcherStateValue -Config $Config -StateKey 'audit_watermark_utc' -StateValue $value)
}

function Get-QCPwDocumentCacheBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string[]]$DocumentGuids
    )
    $cache = @{}
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_DOC_CACHE_SKIPPED' -Message 'Database disabled.' -Data @{ cache = $cache; hits = 0; failedHits = 0 }
    }
    $guids = @($DocumentGuids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($guids.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_DOC_CACHE_NONE' -Message 'No GUIDs.' -Data @{ cache = $cache; hits = 0; failedHits = 0 }
    }
    $hits = 0; $failedHits = 0
    $chunkSize = 100
    for ($i = 0; $i -lt $guids.Count; $i += $chunkSize) {
        $chunk = @($guids[$i..[Math]::Min($i + $chunkSize - 1, $guids.Count - 1)])
        $paramNames = @(); $params = @{}
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $paramNames += "@g$j"
            $params["g$j"] = $chunk[$j]
        }
        $inList = $paramNames -join ','
        $sql = @"
SELECT document_guid, folder_path, description, workflow_state, resolve_failed
FROM pw_document_cache
WHERE document_guid IN ($inList)
  AND expires_at > SYSDATETIMEOFFSET()
"@
        try {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters $params
            if ($res.IsSuccess -and $res.Data.table) {
                foreach ($row in @(_QDB-ConvertDataTableToRowHashtables -Table $res.Data.table)) {
                    $g = ([string]$row.document_guid).Trim().ToLowerInvariant()
                    if ([bool]$row.resolve_failed) {
                        $cache[$g] = @{ resolveFailed = $true }
                        $failedHits++
                    } else {
                        $cache[$g] = @{
                            folderPath = [string]$row.folder_path
                            description = [string]$row.description
                            workflowState = [string]$row.workflow_state
                            resolveFailed = $false
                        }
                        $hits++
                    }
                }
            }
        } catch { }
    }
    return New-QCSuccessResult -Code 'PW_DOC_CACHE_OK' -Message "Cache loaded: $hits hits, $failedHits negative." -Data @{ cache = $cache; hits = $hits; failedHits = $failedHits }
}

function Set-QCPwDocumentCacheEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [string]$FolderPath = '',
        [string]$Description = '',
        [string]$WorkflowState = '',
        [string]$LastAuditAction = '',
        [int]$TtlSeconds = 3600,
        [switch]$ResolveFailed
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }
    $g = ([string]$DocumentGuid).Trim()
    if ([string]::IsNullOrWhiteSpace($g)) { return }
    if ($TtlSeconds -lt 60) { $TtlSeconds = 60 }
    try {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
MERGE pw_document_cache AS tgt
USING (SELECT @guid AS document_guid) AS src
ON tgt.document_guid = src.document_guid
WHEN MATCHED THEN UPDATE SET
    folder_path = CASE WHEN @failed = 1 THEN tgt.folder_path ELSE COALESCE(NULLIF(@folder,''), tgt.folder_path) END,
    description = CASE WHEN @failed = 1 THEN tgt.description ELSE COALESCE(NULLIF(@desc,''), tgt.description) END,
    workflow_state = CASE WHEN @failed = 1 THEN tgt.workflow_state ELSE COALESCE(NULLIF(@state,''), tgt.workflow_state) END,
    resolve_failed = @failed,
    last_audit_action = COALESCE(NULLIF(@action,''), tgt.last_audit_action),
    cached_at = SYSDATETIMEOFFSET(),
    expires_at = DATEADD(SECOND, @ttl, SYSDATETIMEOFFSET())
WHEN NOT MATCHED THEN INSERT
    (document_guid, folder_path, description, workflow_state, resolve_failed, last_audit_action, expires_at)
VALUES
    (@guid, @folder, @desc, @state, @failed, @action, DATEADD(SECOND, @ttl, SYSDATETIMEOFFSET()));
"@ -Parameters @{
            guid = $g; folder = $FolderPath; desc = $Description; state = $WorkflowState
            action = $LastAuditAction; ttl = $TtlSeconds; failed = [bool]$ResolveFailed.IsPresent
        })
    } catch { }
}

function Update-QCProcessingJobCheckpoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][string]$Checkpoint,
        [string]$CheckpointData = ''
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }
    try {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE processing_jobs
SET [checkpoint] = @cp, [checkpoint_data] = @data, last_heartbeat_at = SYSDATETIMEOFFSET()
WHERE job_id = @jobId
"@ -Parameters @{ jobId = $JobId; cp = $Checkpoint; data = $CheckpointData })
    } catch { }
}

function Update-QCProcessingJobHeartbeat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$JobId
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }
    try {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql 'UPDATE processing_jobs SET last_heartbeat_at = SYSDATETIMEOFFSET() WHERE job_id = @jobId' -Parameters @{ jobId = $JobId })
    } catch { }
}

function Get-QCPwFolderGuidCache {
    <#
    .SYNOPSIS
    Loads folder_guid -> folder_path from pw_folder_cache (non-expired, resolved).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    $cache = @{}
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_FOLDER_CACHE_SKIPPED' -Message 'Database disabled.' -Data @{ cache = $cache; count = 0 }
    }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT folder_guid, folder_path
FROM pw_folder_cache
WHERE resolve_failed = 0
  AND NULLIF(LTRIM(RTRIM(folder_path)), '') IS NOT NULL
  AND expires_at > SYSDATETIMEOFFSET()
"@
        if ($res.IsSuccess -and $res.Data -and $res.Data.table) {
            foreach ($row in @(_QDB-ConvertDataTableToRowHashtables -Table $res.Data.table)) {
                $g = [string]$row.folder_guid
                $fp = [string]$row.folder_path
                if ($g -and $fp) { $cache[$g.Trim().ToLowerInvariant()] = $fp }
            }
        }
    } catch { }
    return New-QCSuccessResult -Code 'PW_FOLDER_CACHE_OK' -Message "Loaded $($cache.Count) cached folder GUID paths." -Data @{ cache = $cache; count = $cache.Count }
}

function Get-QCPwFolderCacheBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string[]]$FolderGuids
    )

    $cache = @{}
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_FOLDER_CACHE_BATCH_SKIPPED' -Message 'Database disabled.' -Data @{ cache = $cache; hits = 0; failedHits = 0 }
    }
    $guids = @($FolderGuids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($guids.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_FOLDER_CACHE_BATCH_NONE' -Message 'No folder GUIDs.' -Data @{ cache = $cache; hits = 0; failedHits = 0 }
    }
    $hits = 0; $failedHits = 0
    $chunkSize = 100
    for ($i = 0; $i -lt $guids.Count; $i += $chunkSize) {
        $chunk = @($guids[$i..[Math]::Min($i + $chunkSize - 1, $guids.Count - 1)])
        $paramNames = @(); $params = @{}
        for ($j = 0; $j -lt $chunk.Count; $j++) {
            $paramNames += "@g$j"
            $params["g$j"] = $chunk[$j]
        }
        $inList = $paramNames -join ','
        $sql = @"
SELECT folder_guid, folder_path, watch_root, resolve_failed
FROM pw_folder_cache
WHERE folder_guid IN ($inList)
  AND expires_at > SYSDATETIMEOFFSET()
"@
        try {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters $params
            if ($res.IsSuccess -and $res.Data.table) {
                foreach ($row in @(_QDB-ConvertDataTableToRowHashtables -Table $res.Data.table)) {
                    $g = ([string]$row.folder_guid).Trim().ToLowerInvariant()
                    if ([bool]$row.resolve_failed) {
                        $cache[$g] = @{ resolveFailed = $true }
                        $failedHits++
                    } else {
                        $cache[$g] = @{
                            folderPath = [string]$row.folder_path
                            watchRoot = [string]$row.watch_root
                            resolveFailed = $false
                        }
                        $hits++
                    }
                }
            }
        } catch { }
    }
    return New-QCSuccessResult -Code 'PW_FOLDER_CACHE_BATCH_OK' -Message "Folder cache: $hits hits, $failedHits negative." -Data @{ cache = $cache; hits = $hits; failedHits = $failedHits }
}

function Set-QCPwFolderCacheEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderGuid,
        [string]$FolderPath = '',
        [string]$WatchRoot = '',
        [int]$TtlSeconds = 86400,
        [switch]$ResolveFailed
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }
    $g = ([string]$FolderGuid).Trim().Trim('{}').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($g)) { return }
    if ($TtlSeconds -lt 60) { $TtlSeconds = 60 }
    try {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
MERGE pw_folder_cache AS tgt
USING (SELECT @guid AS folder_guid) AS src
ON tgt.folder_guid = src.folder_guid
WHEN MATCHED THEN UPDATE SET
    folder_path = CASE WHEN @failed = 1 THEN tgt.folder_path ELSE COALESCE(NULLIF(@folder,''), tgt.folder_path) END,
    watch_root = CASE WHEN @failed = 1 THEN tgt.watch_root ELSE COALESCE(NULLIF(@watchRoot,''), tgt.watch_root) END,
    resolve_failed = @failed,
    cached_at = SYSDATETIMEOFFSET(),
    expires_at = DATEADD(SECOND, @ttl, SYSDATETIMEOFFSET())
WHEN NOT MATCHED THEN INSERT
    (folder_guid, folder_path, watch_root, resolve_failed, expires_at)
VALUES
    (@guid, @folder, @watchRoot, @failed, DATEADD(SECOND, @ttl, SYSDATETIMEOFFSET()));
"@ -Parameters @{
            guid = $g; folder = $FolderPath; watchRoot = $WatchRoot; ttl = $TtlSeconds; failed = [bool]$ResolveFailed.IsPresent
        })
    } catch { }
}

Export-ModuleMember -Function Test-QCDatabaseEnabled, Test-QCDatabaseWritesAllowed, Test-QCSheetIndexFolderPath, Get-QCDatabaseConnection, Invoke-QCDatabaseQuery, Invoke-QCDatabaseNonQuery, Invoke-QCDatabaseScalar, Invoke-QCDatabaseBatch, New-QCDatabaseSession, Invoke-QCDatabaseNonQueryWithConnection, Invoke-QCDatabaseScalarWithConnection, Initialize-QCDatabaseSchema, Get-QCProcessingJobType, New-QCStateChangeJobId, Write-QCStateChangeJobTelemetry, Write-QCAuditEventRows, Write-QCJobTelemetry, Write-QCPollRunTelemetry, Write-QCDocumentStateHistoryRow, Write-QCWorkflowEventRow, Write-QCTransitionEvent, Ensure-QCTransitionEvent, Test-QCTransitionEventNotificationSent, Update-QCTransitionEventNotification, Get-QCTransitionEventActor, Get-QCAuditEventActor, Write-QCNotificationTelemetry, Write-QCSheetIndex, Write-QCSheetIndexBatch, Update-QCSheetIndexPwStateName, Update-QCSheetQcPdf, Get-QCPWUnresolvedUserNumbers, Get-QCPWUserIdentity, Write-QCPWUserDirectory, Get-QCDocumentFolderCache, Get-QCUnprocessedAuditEvents, Update-QCAuditEventsResolvedFolders, Mark-QCAuditEventsProcessed, Upsert-QCDocumentActivityFolder, Get-QCWatcherStateValue, Set-QCWatcherStateValue, Get-QCAuditWatermarkUtc, Set-QCAuditWatermarkUtc, Get-QCPwDocumentCacheBatch, Set-QCPwDocumentCacheEntry, Get-QCPwFolderGuidCache, Get-QCPwFolderCacheBatch, Set-QCPwFolderCacheEntry, Update-QCProcessingJobCheckpoint, Update-QCProcessingJobHeartbeat
