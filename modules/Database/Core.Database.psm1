# Core.Database.psm1
# Responsibility: SQL Server connectivity and schema management for QC pipeline telemetry.
# The database is the reporting/control layer. The JSON queue remains the execution source.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Runtime.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Paths.psm1') -Force

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

    $targetVersion = '1.21.0'
    $schemaV1 = _QDB-GetSchemaV1
    $schemaV1_1 = _QDB-GetSchemaV1dot1
    $schemaV1_2 = _QDB-GetSchemaV1dot2
    $schemaV1_3 = _QDB-GetSchemaV1dot3
    $schemaV1_4 = _QDB-GetSchemaV1dot4
    $schemaV1_5 = _QDB-GetSchemaV1dot5
    $schemaV1_6 = _QDB-GetSchemaV1dot6
    $schemaSql = $schemaV1 + [Environment]::NewLine + $schemaV1_1 + [Environment]::NewLine + $schemaV1_2 + [Environment]::NewLine + $schemaV1_3 + [Environment]::NewLine + $schemaV1_4 + [Environment]::NewLine + $schemaV1_5 + [Environment]::NewLine + $schemaV1_6
    $patchSql = (_QDB-GetSchemaV1dot3Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot4Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot5Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot6Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot7Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot8Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot9Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot10Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot11Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot12Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot13Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot14Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot15Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot16Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot17Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot18Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot19Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot20Additive) + [Environment]::NewLine + (_QDB-GetSchemaV1dot21Additive) + [Environment]::NewLine + (_QDB-GetProcessingJobsAdditive)

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
        [void]$insertCmd.Parameters.AddWithValue("@desc", "QC telemetry schema with notification email threading tables")
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
    dedupe_key      NVARCHAR(500),
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

-- qc_comment_runs: one processor execution per lane PDF sync job
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
    dedupe_key      NVARCHAR(500),
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

function _QDB-GetSchemaV1dot10Additive {
    return @'
GO
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL
   AND COL_LENGTH('dbo.processing_jobs', 'dedupe_key') IS NOT NULL
   AND COL_LENGTH('dbo.processing_jobs', 'dedupe_key') < 500
    ALTER TABLE dbo.processing_jobs ALTER COLUMN dedupe_key NVARCHAR(500) NULL;
'@
}

function _QDB-GetSchemaV1dot11Additive {
    return @'
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.sheet_index', 'qc_cycle_id') IS NULL
        ALTER TABLE sheet_index ADD qc_cycle_id NVARCHAR(80) NULL;
    IF COL_LENGTH('dbo.sheet_index', 'qc_cycle_number') IS NULL
        ALTER TABLE sheet_index ADD qc_cycle_number INT NULL;
END
'@
}

function _QDB-GetSchemaV1dot12Additive {
    return @'
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL
   AND COL_LENGTH('dbo.sheet_index', 'qc_cycle_number') IS NOT NULL
   AND EXISTS (
       SELECT 1
       FROM sys.columns c
       INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
       WHERE c.object_id = OBJECT_ID('dbo.sheet_index')
         AND c.name = 'qc_cycle_number'
         AND t.name IN ('int', 'bigint', 'smallint', 'tinyint')
   )
BEGIN
    ALTER TABLE sheet_index ALTER COLUMN qc_cycle_number NVARCHAR(16) NULL;
END
'@
}

function _QDB-GetSchemaV1dot13Additive {
    return @'
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NULL
CREATE TABLE qc_cycle_completions (
    id bigint IDENTITY(1,1) PRIMARY KEY,
    document_guid uniqueidentifier NOT NULL,
    document_name nvarchar(260) NULL,
    qc_cycle_id nvarchar(100) NOT NULL,
    qc_cycle_number int NULL,
    qc_review_type nvarchar(100) NOT NULL,
    completed_at datetime2 NOT NULL,
    completed_by nvarchar(256) NULL,
    transition_event_id bigint NULL,
    audit_event_id bigint NULL,
    created_at datetime2 NOT NULL DEFAULT SYSUTCDATETIME(),
    CONSTRAINT UQ_qc_cycle_completions_cycle UNIQUE (document_guid, qc_cycle_id, qc_review_type)
);
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_cycle_completions_document')
    CREATE INDEX IX_qc_cycle_completions_document ON qc_cycle_completions (document_guid);
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_cycle_completions_completed_at')
    CREATE INDEX IX_qc_cycle_completions_completed_at ON qc_cycle_completions (completed_at);
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.sheet_index', 'production_qc_completed_count') IS NULL
        ALTER TABLE sheet_index ADD production_qc_completed_count int NOT NULL CONSTRAINT DF_sheet_index_production_qc_completed_count DEFAULT 0;
    IF COL_LENGTH('dbo.sheet_index', 'production_qc_last_completed_at') IS NULL
        ALTER TABLE sheet_index ADD production_qc_last_completed_at datetime2 NULL;
    IF COL_LENGTH('dbo.sheet_index', 'peer_review_completed_count') IS NULL
        ALTER TABLE sheet_index ADD peer_review_completed_count int NOT NULL CONSTRAINT DF_sheet_index_peer_review_completed_count DEFAULT 0;
    IF COL_LENGTH('dbo.sheet_index', 'peer_review_last_completed_at') IS NULL
        ALTER TABLE sheet_index ADD peer_review_last_completed_at datetime2 NULL;
    IF COL_LENGTH('dbo.sheet_index', 'independent_check_completed_count') IS NULL
        ALTER TABLE sheet_index ADD independent_check_completed_count int NOT NULL CONSTRAINT DF_sheet_index_independent_check_completed_count DEFAULT 0;
    IF COL_LENGTH('dbo.sheet_index', 'independent_check_last_completed_at') IS NULL
        ALTER TABLE sheet_index ADD independent_check_last_completed_at datetime2 NULL;
END
'@
}

function _QDB-GetSchemaV1dot14Additive {
    return @'
GO
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'transition_event_id') IS NULL
    ALTER TABLE qc_workflow_events ADD transition_event_id INT NULL;
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_transition_event')
    CREATE INDEX IX_qc_workflow_events_transition_event ON qc_workflow_events (transition_event_id);
'@
}

function _QDB-GetSchemaV1dot15Additive {
    return @'
GO
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NULL
CREATE TABLE sheet_packages (
    sheet_package_id                      UNIQUEIDENTIFIER NOT NULL CONSTRAINT DF_sheet_packages_id DEFAULT NEWSEQUENTIALID(),
    sheet_stem                            NVARCHAR(260)    NOT NULL,
    folder_path                           NVARCHAR(1000)   NOT NULL,
    folder_guid                           UNIQUEIDENTIFIER NULL,
    dgn_guid                              UNIQUEIDENTIFIER NULL,
    dgn_name                              NVARCHAR(260)    NULL,
    sheet_pdf_guid                        UNIQUEIDENTIFIER NULL,
    sheet_pdf_name                        NVARCHAR(260)    NULL,
    qc_pdf_guid                           UNIQUEIDENTIFIER NULL,
    qc_pdf_name                           NVARCHAR(260)    NULL,
    pw_state_name                         NVARCHAR(100)    NULL,
    qc_review_type                        NVARCHAR(100)    NULL,
    designer_email                        NVARCHAR(256)    NULL,
    reviewer_email                        NVARCHAR(256)    NULL,
    checker_email                         NVARCHAR(256)    NULL,
    qc_assigned_to                        NVARCHAR(256)    NULL,
    qc_cycle_id                           NVARCHAR(100)    NULL,
    qc_cycle_number                       NVARCHAR(16)     NULL,
    production_qc_completed_count         INT NOT NULL CONSTRAINT DF_sheet_packages_production_qc_completed_count DEFAULT 0,
    production_qc_last_completed_at       DATETIME2        NULL,
    peer_review_completed_count           INT NOT NULL CONSTRAINT DF_sheet_packages_peer_review_completed_count DEFAULT 0,
    peer_review_last_completed_at         DATETIME2        NULL,
    independent_check_completed_count     INT NOT NULL CONSTRAINT DF_sheet_packages_independent_check_completed_count DEFAULT 0,
    independent_check_last_completed_at   DATETIME2        NULL,
    first_seen_at                         DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_sheet_packages_first_seen DEFAULT SYSDATETIMEOFFSET(),
    last_updated_at                       DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_sheet_packages_last_updated DEFAULT SYSDATETIMEOFFSET(),
    CONSTRAINT PK_sheet_packages PRIMARY KEY (sheet_package_id),
    CONSTRAINT UQ_sheet_packages_folder_stem UNIQUE (folder_path, sheet_stem)
);
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_packages_folder')
    CREATE INDEX IX_sheet_packages_folder ON sheet_packages (folder_path);
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_packages_dgn')
    CREATE INDEX IX_sheet_packages_dgn ON sheet_packages (dgn_guid) WHERE dgn_guid IS NOT NULL;
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_packages_sheet_pdf')
    CREATE INDEX IX_sheet_packages_sheet_pdf ON sheet_packages (sheet_pdf_guid) WHERE sheet_pdf_guid IS NOT NULL;
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_packages_qc_pdf')
    CREATE INDEX IX_sheet_packages_qc_pdf ON sheet_packages (qc_pdf_guid) WHERE qc_pdf_guid IS NOT NULL;
GO
IF OBJECT_ID('dbo.sheet_documents', 'U') IS NULL
CREATE TABLE sheet_documents (
    document_guid     UNIQUEIDENTIFIER NOT NULL,
    sheet_package_id  UNIQUEIDENTIFIER NOT NULL,
    document_name     NVARCHAR(260)    NOT NULL,
    document_role     NVARCHAR(50)     NOT NULL,
    pw_state_name     NVARCHAR(100)    NULL,
    extension         NVARCHAR(20)     NULL,
    source_type       NVARCHAR(10)     NULL,
    last_seen_at      DATETIMEOFFSET(3) NULL,
    file_modified_at  DATETIMEOFFSET(3) NULL,
    CONSTRAINT PK_sheet_documents PRIMARY KEY (document_guid),
    CONSTRAINT UQ_sheet_documents_package_role UNIQUE (sheet_package_id, document_role),
    CONSTRAINT FK_sheet_documents_package FOREIGN KEY (sheet_package_id) REFERENCES sheet_packages(sheet_package_id)
);
IF OBJECT_ID('dbo.sheet_documents', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_documents_package')
    CREATE INDEX IX_sheet_documents_package ON sheet_documents (sheet_package_id);
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_index', 'sheet_package_id') IS NULL
    ALTER TABLE sheet_index ADD sheet_package_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_index', 'sheet_package_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_sheet_index_package')
    CREATE INDEX IX_sheet_index_package ON sheet_index (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.document_state_history', 'U') IS NOT NULL AND COL_LENGTH('dbo.document_state_history', 'sheet_package_id') IS NULL
    ALTER TABLE document_state_history ADD sheet_package_id UNIQUEIDENTIFIER NULL;
IF OBJECT_ID('dbo.document_state_history', 'U') IS NOT NULL AND COL_LENGTH('dbo.document_state_history', 'transition_group_id') IS NULL
    ALTER TABLE document_state_history ADD transition_group_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.document_state_history', 'U') IS NOT NULL AND COL_LENGTH('dbo.document_state_history', 'sheet_package_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_state_hist_sheet_package')
    CREATE INDEX IX_state_hist_sheet_package ON document_state_history (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.transition_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.transition_events', 'sheet_package_id') IS NULL
    ALTER TABLE transition_events ADD sheet_package_id UNIQUEIDENTIFIER NULL;
IF OBJECT_ID('dbo.transition_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.transition_events', 'transition_group_id') IS NULL
    ALTER TABLE transition_events ADD transition_group_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.transition_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.transition_events', 'sheet_package_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_transition_sheet_package')
    CREATE INDEX IX_transition_sheet_package ON transition_events (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
IF OBJECT_ID('dbo.transition_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.transition_events', 'transition_group_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_transition_group')
    CREATE INDEX IX_transition_group ON transition_events (transition_group_id) WHERE transition_group_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'sheet_package_id') IS NULL
    ALTER TABLE qc_workflow_events ADD sheet_package_id UNIQUEIDENTIFIER NULL;
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'transition_group_id') IS NULL
    ALTER TABLE qc_workflow_events ADD transition_group_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'sheet_package_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_sheet_package')
    CREATE INDEX IX_qc_workflow_events_sheet_package ON qc_workflow_events (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'transition_group_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_workflow_events_transition_group')
    CREATE INDEX IX_qc_workflow_events_transition_group ON qc_workflow_events (transition_group_id) WHERE transition_group_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'sheet_package_id') IS NULL
    ALTER TABLE processing_jobs ADD sheet_package_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.processing_jobs', 'U') IS NOT NULL AND COL_LENGTH('dbo.processing_jobs', 'sheet_package_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_jobs_sheet_package')
    CREATE INDEX IX_jobs_sheet_package ON processing_jobs (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.notification_log', 'U') IS NOT NULL AND COL_LENGTH('dbo.notification_log', 'sheet_package_id') IS NULL
    ALTER TABLE notification_log ADD sheet_package_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.notification_log', 'U') IS NOT NULL AND COL_LENGTH('dbo.notification_log', 'sheet_package_id') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_notif_sheet_package')
    CREATE INDEX IX_notif_sheet_package ON notification_log (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_sheet_package_status', 'V') IS NOT NULL DROP VIEW v_sheet_package_status;
GO
CREATE VIEW v_sheet_package_status AS
SELECT
    sp.sheet_package_id,
    sp.sheet_stem,
    sp.folder_path,
    sp.folder_guid,
    sp.dgn_guid,
    sp.dgn_name,
    sp.sheet_pdf_guid,
    sp.sheet_pdf_name,
    sp.qc_pdf_guid,
    sp.qc_pdf_name,
    sp.pw_state_name,
    sp.qc_review_type,
    sp.designer_email,
    sp.reviewer_email,
    sp.checker_email,
    sp.qc_assigned_to,
    sp.qc_cycle_id,
    sp.qc_cycle_number,
    sp.production_qc_completed_count,
    sp.production_qc_last_completed_at,
    sp.peer_review_completed_count,
    sp.peer_review_last_completed_at,
    sp.independent_check_completed_count,
    sp.independent_check_last_completed_at,
    sp.first_seen_at,
    sp.last_updated_at
FROM sheet_packages sp;
GO
IF OBJECT_ID('dbo.v_sheet_document_status', 'V') IS NOT NULL DROP VIEW v_sheet_document_status;
GO
CREATE VIEW v_sheet_document_status AS
SELECT
    sd.document_guid,
    sd.sheet_package_id,
    sd.document_name,
    sd.document_role,
    sd.pw_state_name,
    sd.extension,
    sd.source_type,
    sd.last_seen_at,
    sd.file_modified_at,
    sp.sheet_stem,
    sp.folder_path,
    sp.pw_state_name AS package_pw_state_name
FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id;
'@
}

function _QDB-GetSchemaV1dot16Additive {
    return @'
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_cycle_completions', 'sheet_package_id') IS NULL
    ALTER TABLE qc_cycle_completions ADD sheet_package_id UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL
   AND OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_qc_cycle_completions_sheet_package')
    ALTER TABLE qc_cycle_completions ADD CONSTRAINT FK_qc_cycle_completions_sheet_package
        FOREIGN KEY (sheet_package_id) REFERENCES sheet_packages(sheet_package_id);
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_cycle_completions_sheet_package')
    CREATE INDEX IX_qc_cycle_completions_sheet_package ON qc_cycle_completions (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_qc_cycle_completions_package')
    CREATE UNIQUE INDEX UQ_qc_cycle_completions_package ON qc_cycle_completions (sheet_package_id, qc_cycle_id, qc_review_type)
    WHERE sheet_package_id IS NOT NULL;
'@
}

function _QDB-GetSchemaV1dot17Additive {
    return @'
GO
IF OBJECT_ID('dbo.v_sheet_package_status', 'V') IS NOT NULL DROP VIEW v_sheet_package_status;
GO
CREATE VIEW v_sheet_package_status AS
SELECT
    sp.sheet_package_id,
    sp.sheet_stem,
    sp.folder_path,
    sp.pw_state_name,
    sp.qc_review_type,
    sp.qc_assigned_to,
    sp.production_qc_completed_count,
    sp.peer_review_completed_count,
    sp.independent_check_completed_count,
    sp.production_qc_last_completed_at,
    sp.peer_review_last_completed_at,
    sp.independent_check_last_completed_at,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid
FROM sheet_packages sp;
GO
IF OBJECT_ID('dbo.v_sheet_package_cycle_aging', 'V') IS NOT NULL DROP VIEW v_sheet_package_cycle_aging;
GO
CREATE VIEW v_sheet_package_cycle_aging AS
SELECT
    sp.sheet_package_id,
    sp.sheet_stem,
    sp.pw_state_name AS current_state,
    sp.qc_cycle_id AS current_cycle_id,
    sp.qc_cycle_number AS current_cycle_number,
    CASE
        WHEN sh.last_state_change_at IS NOT NULL
        THEN DATEDIFF(day, CAST(sh.last_state_change_at AS datetimeoffset), SYSDATETIMEOFFSET())
        ELSE DATEDIFF(day, sp.last_updated_at, SYSDATETIMEOFFSET())
    END AS days_in_current_state,
    CASE
        WHEN lc.last_completion_at IS NOT NULL
        THEN DATEDIFF(day, lc.last_completion_at, CAST(SYSDATETIMEOFFSET() AS datetime2))
        ELSE NULL
    END AS days_since_last_completion,
    sp.production_qc_completed_count AS production_completion_count,
    sp.peer_review_completed_count AS peer_completion_count,
    sp.independent_check_completed_count AS independent_completion_count
FROM sheet_packages sp
OUTER APPLY (
    SELECT MAX(dsh.captured_at) AS last_state_change_at
    FROM document_state_history dsh
    WHERE dsh.sheet_package_id = sp.sheet_package_id
      AND dsh.field_name = 'pw_state_name'
      AND ISNULL(dsh.new_value, '') = ISNULL(sp.pw_state_name, '')
) sh
OUTER APPLY (
    SELECT MAX(v) AS last_completion_at
    FROM (VALUES
        (sp.production_qc_last_completed_at),
        (sp.peer_review_last_completed_at),
        (sp.independent_check_last_completed_at)
    ) AS t(v)
    WHERE v IS NOT NULL
) lc;
'@
}

function _QDB-GetSchemaV1dot18Additive {
    return @'
GO
IF OBJECT_ID('dbo.sheet_index', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_index', 'qc_process_type') IS NULL
    ALTER TABLE sheet_index ADD qc_process_type NVARCHAR(32) NULL;
GO
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_packages', 'qc_process_type') IS NULL
    ALTER TABLE sheet_packages ADD qc_process_type NVARCHAR(32) NULL;
GO
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_packages', 'qc_chk_pdf_guid') IS NULL
    ALTER TABLE sheet_packages ADD qc_chk_pdf_guid UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_packages', 'qc_chk_pdf_name') IS NULL
    ALTER TABLE sheet_packages ADD qc_chk_pdf_name NVARCHAR(512) NULL;
GO
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_packages', 'qc_rev_pdf_guid') IS NULL
    ALTER TABLE sheet_packages ADD qc_rev_pdf_guid UNIQUEIDENTIFIER NULL;
GO
IF OBJECT_ID('dbo.sheet_packages', 'U') IS NOT NULL AND COL_LENGTH('dbo.sheet_packages', 'qc_rev_pdf_name') IS NULL
    ALTER TABLE sheet_packages ADD qc_rev_pdf_name NVARCHAR(512) NULL;
GO
IF OBJECT_ID('dbo.qc_workflow_events', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_workflow_events', 'qc_process_type') IS NULL
    ALTER TABLE qc_workflow_events ADD qc_process_type NVARCHAR(32) NULL;
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND COL_LENGTH('dbo.qc_cycle_completions', 'qc_process_type') IS NULL
    ALTER TABLE qc_cycle_completions ADD qc_process_type NVARCHAR(32) NULL;
GO
IF OBJECT_ID('dbo.notification_log', 'U') IS NOT NULL AND COL_LENGTH('dbo.notification_log', 'qc_process_type') IS NULL
    ALTER TABLE notification_log ADD qc_process_type NVARCHAR(32) NULL;
GO
IF OBJECT_ID('dbo.qc_cycle_completions', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_qc_cycle_completions_package_process')
    CREATE UNIQUE INDEX UQ_qc_cycle_completions_package_process ON qc_cycle_completions (sheet_package_id, qc_cycle_id, qc_process_type)
    WHERE sheet_package_id IS NOT NULL AND qc_process_type IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_sheet_package_status', 'V') IS NOT NULL DROP VIEW v_sheet_package_status;
GO
CREATE VIEW v_sheet_package_status AS
SELECT
    sp.sheet_package_id,
    sp.sheet_stem,
    sp.folder_path,
    sp.pw_state_name,
    sp.qc_review_type,
    COALESCE(
        sp.qc_process_type,
        CASE LOWER(LTRIM(RTRIM(ISNULL(sp.qc_review_type, ''))))
            WHEN 'production qc' THEN 'production'
            WHEN 'independent check' THEN 'check'
            WHEN 'peer review' THEN 'review'
            WHEN 'peer_review' THEN 'review'
            WHEN 'independent_check' THEN 'check'
            ELSE LOWER(LTRIM(RTRIM(sp.qc_review_type)))
        END
    ) AS qc_process_type,
  CASE COALESCE(
        sp.qc_process_type,
        CASE LOWER(LTRIM(RTRIM(ISNULL(sp.qc_review_type, ''))))
            WHEN 'production qc' THEN 'production'
            WHEN 'independent check' THEN 'check'
            WHEN 'peer review' THEN 'review'
            WHEN 'peer_review' THEN 'review'
            WHEN 'independent_check' THEN 'check'
            ELSE LOWER(LTRIM(RTRIM(sp.qc_review_type)))
        END
    )
        WHEN 'production' THEN 'Production'
        WHEN 'check' THEN 'Check'
        WHEN 'review' THEN 'Review'
        ELSE NULL
    END AS qc_process_type_display,
    sp.qc_assigned_to,
    sp.production_qc_completed_count,
    sp.peer_review_completed_count AS review_completed_count,
    sp.independent_check_completed_count AS check_completed_count,
    sp.production_qc_last_completed_at,
    sp.peer_review_last_completed_at AS review_last_completed_at,
    sp.independent_check_last_completed_at AS check_last_completed_at,
    sp.peer_review_completed_count,
    sp.independent_check_completed_count,
    sp.peer_review_last_completed_at,
    sp.independent_check_last_completed_at,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid,
    sp.qc_chk_pdf_guid,
    sp.qc_rev_pdf_guid
FROM sheet_packages sp;
'@
}

function _QDB-GetSchemaV1dot19Additive {
    return @'
GO
IF OBJECT_ID('dbo.sheet_package_qc_pdfs', 'U') IS NULL
CREATE TABLE sheet_package_qc_pdfs (
    id                        INT IDENTITY(1,1) NOT NULL,
    sheet_package_id          UNIQUEIDENTIFIER NOT NULL,
    qc_process_type           NVARCHAR(32) NOT NULL,
    document_guid             UNIQUEIDENTIFIER NOT NULL,
    document_name             NVARCHAR(512) NOT NULL,
    folder_path               NVARCHAR(1000) NULL,
    current_pw_state          NVARCHAR(256) NULL,
    previous_pw_state         NVARCHAR(256) NULL,
    assigned_reviewer_email   NVARCHAR(320) NULL,
    assigned_reviewer_name    NVARCHAR(256) NULL,
    review_cycle              INT NULL,
    created_at                DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_sheet_package_qc_pdfs_created DEFAULT SYSDATETIMEOFFSET(),
    updated_at                DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_sheet_package_qc_pdfs_updated DEFAULT SYSDATETIMEOFFSET(),
    last_seen_at              DATETIMEOFFSET(3) NULL,
    is_active                 BIT NOT NULL CONSTRAINT DF_sheet_package_qc_pdfs_active DEFAULT 1,
    CONSTRAINT PK_sheet_package_qc_pdfs PRIMARY KEY (id),
    CONSTRAINT FK_sheet_package_qc_pdfs_package FOREIGN KEY (sheet_package_id) REFERENCES sheet_packages(sheet_package_id),
    CONSTRAINT CK_sheet_package_qc_pdfs_process CHECK (qc_process_type IN ('production', 'check', 'review'))
);
GO
IF OBJECT_ID('dbo.sheet_package_qc_pdfs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'ux_sheet_package_qc_pdfs_package_process_active')
    CREATE UNIQUE INDEX ux_sheet_package_qc_pdfs_package_process_active
    ON sheet_package_qc_pdfs(sheet_package_id, qc_process_type)
    WHERE is_active = 1;
GO
IF OBJECT_ID('dbo.sheet_package_qc_pdfs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'ix_sheet_package_qc_pdfs_document_guid')
    CREATE INDEX ix_sheet_package_qc_pdfs_document_guid ON sheet_package_qc_pdfs(document_guid);
GO
IF OBJECT_ID('dbo.sheet_package_qc_pdfs', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'ix_sheet_package_qc_pdfs_sheet_package_id')
    CREATE INDEX ix_sheet_package_qc_pdfs_sheet_package_id ON sheet_package_qc_pdfs(sheet_package_id);
GO
IF OBJECT_ID('dbo.v_sheet_package_qc_pdf_matrix', 'V') IS NOT NULL DROP VIEW v_sheet_package_qc_pdf_matrix;
GO
CREATE VIEW v_sheet_package_qc_pdf_matrix AS
SELECT
    sp.sheet_package_id,
    sp.sheet_stem AS package_key,
    sp.sheet_stem AS sheet_number,
    sp.folder_path,
    MAX(CASE WHEN q.qc_process_type = 'production' THEN CAST(q.document_guid AS NVARCHAR(36)) END) AS production_qc_pdf_guid,
    MAX(CASE WHEN q.qc_process_type = 'production' THEN q.document_name END) AS production_qc_pdf_name,
    MAX(CASE WHEN q.qc_process_type = 'production' THEN q.current_pw_state END) AS production_state,
    MAX(CASE WHEN q.qc_process_type = 'check' THEN CAST(q.document_guid AS NVARCHAR(36)) END) AS check_qc_pdf_guid,
    MAX(CASE WHEN q.qc_process_type = 'check' THEN q.document_name END) AS check_qc_pdf_name,
    MAX(CASE WHEN q.qc_process_type = 'check' THEN q.current_pw_state END) AS check_state,
    MAX(CASE WHEN q.qc_process_type = 'review' THEN CAST(q.document_guid AS NVARCHAR(36)) END) AS review_qc_pdf_guid,
    MAX(CASE WHEN q.qc_process_type = 'review' THEN q.document_name END) AS review_qc_pdf_name,
    MAX(CASE WHEN q.qc_process_type = 'review' THEN q.current_pw_state END) AS review_state
FROM sheet_packages sp
LEFT JOIN sheet_package_qc_pdfs q
    ON q.sheet_package_id = sp.sheet_package_id
    AND q.is_active = 1
GROUP BY
    sp.sheet_package_id,
    sp.sheet_stem,
    sp.folder_path;
GO
IF OBJECT_ID('dbo.v_sheet_package_status', 'V') IS NOT NULL DROP VIEW v_sheet_package_status;
GO
CREATE VIEW v_sheet_package_status AS
SELECT
    sp.sheet_package_id,
    sp.sheet_stem,
    sp.folder_path,
    sp.qc_review_type,
    COALESCE(
        sp.qc_process_type,
        CASE LOWER(LTRIM(RTRIM(ISNULL(sp.qc_review_type, ''))))
            WHEN 'production qc' THEN 'production'
            WHEN 'independent check' THEN 'check'
            WHEN 'peer review' THEN 'review'
            WHEN 'peer_review' THEN 'review'
            WHEN 'independent_check' THEN 'check'
            ELSE LOWER(LTRIM(RTRIM(sp.qc_review_type)))
        END
    ) AS qc_process_type,
    CASE COALESCE(
        sp.qc_process_type,
        CASE LOWER(LTRIM(RTRIM(ISNULL(sp.qc_review_type, ''))))
            WHEN 'production qc' THEN 'production'
            WHEN 'independent check' THEN 'check'
            WHEN 'peer review' THEN 'review'
            WHEN 'peer_review' THEN 'review'
            WHEN 'independent_check' THEN 'check'
            ELSE LOWER(LTRIM(RTRIM(sp.qc_review_type)))
        END
    )
        WHEN 'production' THEN 'Production'
        WHEN 'check' THEN 'Check'
        WHEN 'review' THEN 'Review'
        ELSE NULL
    END AS qc_process_type_display,
    sp.qc_assigned_to,
    sp.production_qc_completed_count,
    sp.peer_review_completed_count AS review_completed_count,
    sp.independent_check_completed_count AS check_completed_count,
    sp.production_qc_last_completed_at,
    sp.peer_review_last_completed_at AS review_last_completed_at,
    sp.independent_check_last_completed_at AS check_last_completed_at,
    sp.peer_review_completed_count,
    sp.independent_check_completed_count,
    sp.peer_review_last_completed_at,
    sp.independent_check_last_completed_at,
    CASE WHEN sp.production_qc_completed_count > 0 THEN 1 ELSE 0 END AS production_completed,
    CASE WHEN sp.independent_check_completed_count > 0 THEN 1 ELSE 0 END AS check_completed,
    CASE WHEN sp.peer_review_completed_count > 0 THEN 1 ELSE 0 END AS review_completed,
    m.production_state,
    m.check_state,
    m.review_state,
    sp.dgn_guid,
    sp.sheet_pdf_guid,
    sp.qc_pdf_guid,
    sp.qc_chk_pdf_guid,
    sp.qc_rev_pdf_guid
FROM sheet_packages sp
LEFT JOIN v_sheet_package_qc_pdf_matrix m ON m.sheet_package_id = sp.sheet_package_id;
'@
}

function _QDB-GetSchemaV1dot20Additive {
    return @'
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NULL
CREATE TABLE automation_events (
    id                BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    ts                DATETIMEOFFSET(3) NOT NULL,
    process_name      NVARCHAR(128) NOT NULL,
    run_id            NVARCHAR(64) NULL,
    level             NVARCHAR(32) NOT NULL,
    code              NVARCHAR(128) NOT NULL,
    message           NVARCHAR(2000) NULL,
    job_id            NVARCHAR(256) NULL,
    document_guid     NVARCHAR(36) NULL,
    sheet_package_id  NVARCHAR(36) NULL,
    audit_event_id    BIGINT NULL,
    folder_path       NVARCHAR(1000) NULL,
    data_json         NVARCHAR(MAX) NOT NULL,
    dedupe_key        NVARCHAR(128) NULL,
    created_at        DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_automation_events_created DEFAULT SYSDATETIMEOFFSET()
);
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_ts')
    CREATE INDEX IX_automation_events_ts ON automation_events (ts DESC);
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_code')
    CREATE INDEX IX_automation_events_code ON automation_events (code);
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_process_name')
    CREATE INDEX IX_automation_events_process_name ON automation_events (process_name);
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_job_id')
    CREATE INDEX IX_automation_events_job_id ON automation_events (job_id) WHERE job_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_document_guid')
    CREATE INDEX IX_automation_events_document_guid ON automation_events (document_guid) WHERE document_guid IS NOT NULL;
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_sheet_package_id')
    CREATE INDEX IX_automation_events_sheet_package_id ON automation_events (sheet_package_id) WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_audit_event_id')
    CREATE INDEX IX_automation_events_audit_event_id ON automation_events (audit_event_id) WHERE audit_event_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_automation_events_level')
    CREATE INDEX IX_automation_events_level ON automation_events (level);
GO
IF OBJECT_ID('dbo.automation_events', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_automation_events_dedupe')
    CREATE UNIQUE INDEX UX_automation_events_dedupe ON automation_events (dedupe_key) WHERE dedupe_key IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_mcp_automation_events_recent', 'V') IS NOT NULL DROP VIEW v_mcp_automation_events_recent;
GO
CREATE VIEW v_mcp_automation_events_recent AS
SELECT TOP (100000) *
FROM automation_events
WHERE ts >= DATEADD(day, -14, SYSDATETIMEOFFSET())
ORDER BY ts DESC;
GO
IF OBJECT_ID('dbo.v_mcp_process_health', 'V') IS NOT NULL DROP VIEW v_mcp_process_health;
GO
CREATE VIEW v_mcp_process_health AS
SELECT
    process_name,
    MAX(ts) AS last_event_at,
    SUM(CASE WHEN level IN ('Error', 'error') THEN 1 ELSE 0 END) AS error_count_24h,
    SUM(CASE WHEN level IN ('Warning', 'warning') THEN 1 ELSE 0 END) AS warning_count_24h,
    SUM(CASE WHEN code = 'WORKER_NO_JOB' THEN 1 ELSE 0 END) AS no_job_count_24h,
    COUNT(*) AS event_count_24h
FROM automation_events
WHERE ts >= DATEADD(hour, -24, SYSDATETIMEOFFSET())
GROUP BY process_name;
GO
IF OBJECT_ID('dbo.v_mcp_job_timeline', 'V') IS NOT NULL DROP VIEW v_mcp_job_timeline;
GO
CREATE VIEW v_mcp_job_timeline AS
SELECT id, ts, process_name, run_id, level, code, message, job_id, document_guid, sheet_package_id, folder_path, data_json
FROM automation_events
WHERE job_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_mcp_document_debug_events', 'V') IS NOT NULL DROP VIEW v_mcp_document_debug_events;
GO
CREATE VIEW v_mcp_document_debug_events AS
SELECT id, ts, process_name, run_id, level, code, message, job_id, document_guid, sheet_package_id, folder_path, data_json
FROM automation_events
WHERE document_guid IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_mcp_package_debug_events', 'V') IS NOT NULL DROP VIEW v_mcp_package_debug_events;
GO
CREATE VIEW v_mcp_package_debug_events AS
SELECT id, ts, process_name, run_id, level, code, message, job_id, document_guid, sheet_package_id, folder_path, data_json
FROM automation_events
WHERE sheet_package_id IS NOT NULL;
GO
IF OBJECT_ID('dbo.v_mcp_audit_scan_history', 'V') IS NOT NULL DROP VIEW v_mcp_audit_scan_history;
GO
CREATE VIEW v_mcp_audit_scan_history AS
SELECT id, ts, process_name, run_id, level, code, message, job_id, document_guid, sheet_package_id, audit_event_id, folder_path, data_json
FROM automation_events
WHERE code LIKE 'WATCH_AUDIT_%'
   OR code LIKE 'AUDIT_EVENTS_%'
   OR code LIKE 'AUDIT_%';
GO
IF OBJECT_ID('dbo.v_mcp_recent_errors', 'V') IS NOT NULL DROP VIEW v_mcp_recent_errors;
GO
CREATE VIEW v_mcp_recent_errors AS
SELECT id, ts, process_name, run_id, level, code, message, job_id, document_guid, sheet_package_id, folder_path, data_json
FROM automation_events
WHERE level IN ('Error', 'error', 'Warning', 'warning')
  AND ts >= DATEADD(day, -7, SYSDATETIMEOFFSET());
'@
}

function _QDB-GetSchemaV1dot21Additive {
    return @'
GO
IF OBJECT_ID('dbo.qc_notification_threads', 'U') IS NULL
CREATE TABLE qc_notification_threads (
    id                                INT IDENTITY(1,1) NOT NULL,
    sheet_package_id                  UNIQUEIDENTIFIER NOT NULL,
    review_type                       NVARCHAR(128) NOT NULL,
    status                            NVARCHAR(32) NOT NULL CONSTRAINT DF_qc_notification_threads_status DEFAULT 'active',
    graph_conversation_id             NVARCHAR(256) NULL,
    latest_graph_message_id           NVARCHAR(256) NULL,
    latest_graph_immutable_message_id NVARCHAR(256) NULL,
    latest_internet_message_id        NVARCHAR(512) NULL,
    created_at                        DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_qc_notification_threads_created DEFAULT SYSDATETIMEOFFSET(),
    updated_at                        DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_qc_notification_threads_updated DEFAULT SYSDATETIMEOFFSET(),
    superseded_at                     DATETIMEOFFSET(3) NULL,
    CONSTRAINT PK_qc_notification_threads PRIMARY KEY (id),
    CONSTRAINT CK_qc_notification_threads_status CHECK (status IN ('active', 'superseded', 'invalid'))
);
GO
IF OBJECT_ID('dbo.qc_notification_threads', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UX_qc_notification_threads_active_package_review')
    CREATE UNIQUE INDEX UX_qc_notification_threads_active_package_review
    ON qc_notification_threads (sheet_package_id, review_type)
    WHERE status = 'active';
GO
IF OBJECT_ID('dbo.qc_notification_threads', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_notification_threads_package')
    CREATE INDEX IX_qc_notification_threads_package ON qc_notification_threads (sheet_package_id);
GO
IF OBJECT_ID('dbo.qc_notification_messages', 'U') IS NULL
CREATE TABLE qc_notification_messages (
    id                          INT IDENTITY(1,1) NOT NULL,
    thread_id                   INT NOT NULL,
    notification_log_id         INT NULL,
    workflow_event              NVARCHAR(100) NULL,
    graph_message_id            NVARCHAR(256) NULL,
    graph_immutable_message_id  NVARCHAR(256) NULL,
    graph_conversation_id       NVARCHAR(256) NULL,
    internet_message_id         NVARCHAR(512) NULL,
    send_mode                   NVARCHAR(32) NOT NULL,
    delivery_status             NVARCHAR(32) NOT NULL CONSTRAINT DF_qc_notification_messages_delivery DEFAULT 'sent',
    sent_at                     DATETIMEOFFSET(3) NULL,
    error_message               NVARCHAR(2000) NULL,
    created_at                  DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_qc_notification_messages_created DEFAULT SYSDATETIMEOFFSET(),
    updated_at                  DATETIMEOFFSET(3) NOT NULL CONSTRAINT DF_qc_notification_messages_updated DEFAULT SYSDATETIMEOFFSET(),
    CONSTRAINT PK_qc_notification_messages PRIMARY KEY (id),
    CONSTRAINT FK_qc_notification_messages_thread FOREIGN KEY (thread_id) REFERENCES qc_notification_threads(id),
    CONSTRAINT CK_qc_notification_messages_send_mode CHECK (send_mode IN ('root', 'reply', 'replacement_root', 'unthreaded')),
    CONSTRAINT CK_qc_notification_messages_delivery CHECK (delivery_status IN ('sent', 'failed', 'pending'))
);
GO
IF OBJECT_ID('dbo.qc_notification_messages', 'U') IS NOT NULL
   AND OBJECT_ID('dbo.notification_log', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_qc_notification_messages_log')
    ALTER TABLE qc_notification_messages
    ADD CONSTRAINT FK_qc_notification_messages_log FOREIGN KEY (notification_log_id) REFERENCES notification_log(id);
GO
IF OBJECT_ID('dbo.qc_notification_messages', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_notification_messages_thread')
    CREATE INDEX IX_qc_notification_messages_thread ON qc_notification_messages (thread_id);
GO
IF OBJECT_ID('dbo.qc_notification_messages', 'U') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_qc_notification_messages_notification_log')
    CREATE INDEX IX_qc_notification_messages_notification_log ON qc_notification_messages (notification_log_id) WHERE notification_log_id IS NOT NULL;
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
        [AllowEmptyString()][string]$PreviousState = '',
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
        -ResultData $resultJson -DocumentGuid $DocumentGuid
}

function _QDB-IsSheetScopedProcessingJobType {
    param([Parameter(Mandatory)][string]$JobType)
    $t = ([string]$JobType).Trim()
    if ($t -match '^(QC_PREPEND|QC_FINALIZE|QC_STATE|QC_NOTIFICATION)$') { return $true }
    if ($t -match '^QC_COMMENT') { return $true }
    return $false
}

function _QDB-ExtractDocumentGuidFromTelemetryJson {
    param([string]$JsonText)
    if ([string]::IsNullOrWhiteSpace($JsonText)) { return '' }
    try {
        $obj = $JsonText | ConvertFrom-Json
        foreach ($name in @('documentGuid', 'document_guid', 'triggerDocumentGuid', 'qcPdfGuid')) {
            try {
                if ($obj.PSObject.Properties[$name] -and -not [string]::IsNullOrWhiteSpace([string]$obj.$name)) {
                    return ([string]$obj.$name).Trim()
                }
            } catch { }
        }
    } catch { }
    return ''
}

function _QDB-ResolveSheetPackageIdForJobTelemetry {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$JobType,
        [string]$DocumentGuid = '',
        [string]$SourcePath = '',
        [string]$SourceFolder = '',
        [string]$ResultData = '',
        [Nullable[guid]]$SheetPackageId = $null
    )
    if (-not (_QDB-IsSheetScopedProcessingJobType -JobType $JobType)) { return $null }
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) { return $SheetPackageId }

    $docGuid = ([string]$DocumentGuid).Trim()
    if ([string]::IsNullOrWhiteSpace($docGuid)) {
        $docGuid = _QDB-ExtractDocumentGuidFromTelemetryJson -JsonText $ResultData
    }
    if ($docGuid) {
        $pkg = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $docGuid
        if ($pkg) { return $pkg }
    }

    $folder = _QDB-NormalizeTelemetryPath -Path $SourceFolder
    $docName = ''
    if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
        $docName = [System.IO.Path]::GetFileName(([string]$SourcePath).Trim())
    }
    if ($folder -and $docName) {
        $resolved = Resolve-SheetPackageFromDocument -DocumentGuid $docGuid -DocumentName $docName -FolderPath $folder
        if ($resolved.isSheetPackageMember) {
            return (Resolve-SheetPackageIdForSheetGroup -Config $Config -FolderPath $folder `
                -SheetStem $resolved.sheetStem -DocumentGuid $docGuid -DocumentName $docName)
        }
    }
    return $null
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
        [string]$ResultData,
        [string]$DocumentGuid = '',
        [Nullable[guid]]$SheetPackageId = $null
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
    $resolvedPackageId = _QDB-ResolveSheetPackageIdForJobTelemetry -Config $Config -JobType $telemetryJobType `
        -DocumentGuid $DocumentGuid -SourcePath $SourcePath -SourceFolder $SourceFolder `
        -ResultData $resultPayload -SheetPackageId $SheetPackageId
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
    result_data = @resultData,
    sheet_package_id = COALESCE(@sheetPackageId, tgt.sheet_package_id)
WHEN NOT MATCHED THEN INSERT
    (job_id, job_type, status, source_path, source_folder, dedupe_key, trigger_source, started_at, attempt_count, duration_ms, error_code, error_message, result_data, sheet_package_id)
VALUES
    (@jobId, @jobType, @status, @sourcePath, @sourceFolder, @dedupeKey, @triggerSource, @startedAt, @attemptCount, @durationMs, @errorCode, @errorMessage, @resultData, @sheetPackageId);
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
            sheetPackageId = if ($null -ne $resolvedPackageId) { $resolvedPackageId } else { [DBNull]::Value }
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
        [Nullable[long]]$SourceAuditId = $null,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
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
    (document_guid, document_name, folder_path, event_type, source_audit_id, old_value, new_value, field_name, changed_by_user, changed_by_username, sheet_package_id, transition_group_id)
VALUES
    (@documentGuid, @documentName, @folderPath, @eventType, @sourceAuditId, @oldValue, @newValue, @fieldName, @changedByUser, @changedByUsername, @sheetPackageId, @transitionGroupId)
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
            sheetPackageId = if ($null -ne $SheetPackageId) { $SheetPackageId } else { [DBNull]::Value }
            transitionGroupId = if ($null -ne $TransitionGroupId) { $TransitionGroupId } else { [DBNull]::Value }
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
        [Nullable[int]]$TransitionEventId = $null,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null,
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
    (run_id, job_id, document_id, event_type, previous_pw_state, target_pw_state, decision_code, processor_version, qc_review_type, payload_json, transition_event_id, sheet_package_id, transition_group_id)
VALUES
    (@runId, @jobId, @documentId, @eventType, @prev, @target, @decisionCode, @procVer, @qcReviewType, @payload, @transitionEventId, @sheetPackageId, @transitionGroupId)
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
            transitionEventId = if ($null -ne $TransitionEventId -and $TransitionEventId -gt 0) { [int]$TransitionEventId } else { [DBNull]::Value }
            sheetPackageId = if ($null -ne $SheetPackageId) { $SheetPackageId } else { [DBNull]::Value }
            transitionGroupId = if ($null -ne $TransitionGroupId) { $TransitionGroupId } else { [DBNull]::Value }
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
        [string]$ChangedByUsername = '',
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
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
        -ChangedByUsername $ChangedByUsername -SheetPackageId $SheetPackageId -TransitionGroupId $TransitionGroupId
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
        [string]$ChangedByUsername = '',
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
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
    (document_guid, document_name, folder_path, transition_type, from_value, to_value, trigger_audit_id, job_id, job_type, changed_by_user, changed_by_username, sheet_package_id, transition_group_id)
OUTPUT INSERTED.id
VALUES
    (@documentGuid, @documentName, @folderPath, @transitionType, @fromValue, @toValue, @triggerAuditId, @jobId, @jobType, @changedByUser, @changedByUsername, @sheetPackageId, @transitionGroupId)
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
            sheetPackageId = if ($null -ne $SheetPackageId) { $SheetPackageId } else { [DBNull]::Value }
            transitionGroupId = if ($null -ne $TransitionGroupId) { $TransitionGroupId } else { [DBNull]::Value }
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
        [Nullable[int]]$TransitionId,
        [Nullable[guid]]$SheetPackageId = $null
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return New-QCSuccessResult -Code 'NOTIFICATION_TELEMETRY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false } }
    $FolderPath = _QDB-NormalizeTelemetryPath -Path $FolderPath
    $resolvedPackageId = $SheetPackageId
    if ($null -eq $resolvedPackageId -and -not [string]::IsNullOrWhiteSpace($DocumentGuid)) {
        $resolvedPackageId = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $DocumentGuid
    }
    try {
        $sql = @"
INSERT INTO notification_log
    (event_type, document_guid, document_name, folder_path, recipients, subject, dedupe_key, provider, success, error_message, transition_id, sheet_package_id)
OUTPUT INSERTED.id
VALUES
    (@eventType, @documentGuid, @documentName, @folderPath, @recipients, @subject, @dedupeKey, @provider, @success, @errorMessage, @transitionId, @sheetPackageId)
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
            sheetPackageId = if ($null -ne $resolvedPackageId) { $resolvedPackageId } else { [DBNull]::Value }
        }
        $dbRes = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            Write-QCJsonLog -Flush -Level 'Error' -Code 'NOTIFICATION_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='insert_notification_log'; eventType=$EventType }
            return New-QCErrorResult -Code 'NOTIFICATION_TELEMETRY_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation='insert_notification_log'; eventType=$EventType }
        }
        $notificationLogId = $null
        if ($null -ne $dbRes.Data.value) {
            try { $notificationLogId = [int]$dbRes.Data.value } catch { $notificationLogId = $null }
        }
        return New-QCSuccessResult -Code 'NOTIFICATION_TELEMETRY_WRITTEN' -Message 'Notification telemetry inserted.' -Data @{ written = $true; notificationLogId = $notificationLogId }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'NOTIFICATION_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='insert_notification_log'; eventType=$EventType }
        return New-QCErrorResult -Code 'NOTIFICATION_TELEMETRY_EXCEPTION' -Message $msg -Data @{ operation='insert_notification_log'; eventType=$EventType }
    }
}

function _QDB-GetSheetStemFromDocumentName {
    param([Parameter(Mandatory)][string]$DocumentName)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ([string]::IsNullOrWhiteSpace($stem)) { return '' }
    if ($stem -match '(?i)-(prod|chk|rev)$') { $stem = $stem -replace '(?i)-(prod|chk|rev)$', '' }
    return $stem
}

function _QDB-ResolveSheetDocumentRole {
    param([string]$DocumentName)
    $dn = [string]$DocumentName
    if ($dn -match '(?i)-(prod|chk|rev)\.pdf$') { return 'qc_pdf' }
    if ($dn -match '(?i)\.dgn$') { return 'dgn' }
    if ($dn -match '(?i)\.pdf$') { return 'sheet_pdf' }
    return 'other'
}

function _QDB-GetQcPdfLaneFromDocumentName {
    param([string]$DocumentName)
    $dn = [string]$DocumentName
    if ($dn -match '(?i)-prod\.pdf$') { return 'production' }
    if ($dn -match '(?i)-chk\.pdf$') { return 'check' }
    if ($dn -match '(?i)-rev\.pdf$') { return 'review' }
    return ''
}

function _QDB-TryParseDocumentGuid {
    param([string]$DocumentGuid)
    $g = ([string]$DocumentGuid).Trim()
    if ($g.Length -lt 5) { return $null }
    try {
        return [guid]::Parse($g)
    } catch {
        return $null
    }
}

function Resolve-SheetPackageFromDocument {
    <#
    .SYNOPSIS
    Resolves sheet stem and document role from a ProjectWise document identity.
    #>
    [CmdletBinding()]
    param(
        [string]$DocumentGuid = '',
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$FolderPath = ''
    )
    $stem = _QDB-GetSheetStemFromDocumentName -DocumentName $DocumentName
    $role = _QDB-ResolveSheetDocumentRole -DocumentName $DocumentName
    return @{
        documentGuid = ([string]$DocumentGuid).Trim()
        documentName = ([string]$DocumentName).Trim()
        folderPath = ([string]$FolderPath).Trim()
        sheetStem = $stem
        documentRole = $role
        isSheetPackageMember = ($role -in @('dgn', 'sheet_pdf', 'qc_pdf'))
    }
}

function Get-SheetPackageIdForDocument {
    <#
    .SYNOPSIS
    Returns sheet_package_id for a physical document GUID, if known.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return $null }
    $parsed = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    if (-not $parsed) { return $null }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sheet_package_id
FROM sheet_documents
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $parsed }
        if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
            $val = $res.Data.table.Rows[0].sheet_package_id
            if ($val -isnot [DBNull]) { return [guid]$val }
        }
    } catch { }
    return $null
}

function Resolve-SheetPackageIdForSheetGroup {
    <#
    .SYNOPSIS
    Resolves sheet_package_id for a logical sheet group from document identity or folder/stem.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SheetStem = '',
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [Nullable[guid]]$SheetPackageId = $null
    )
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) { return $SheetPackageId }
    if (-not [string]::IsNullOrWhiteSpace($DocumentGuid)) {
        $pkg = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $DocumentGuid
        if ($pkg) { return $pkg }
    }
    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath
    $stem = ([string]$SheetStem).Trim()
    if ($folder -and $stem -and (_QDB-IsEnabled -Config $Config)) {
        try {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sheet_package_id
FROM sheet_packages
WHERE folder_path = @folderPath AND sheet_stem = @sheetStem
"@ -Parameters @{ folderPath = $folder; sheetStem = $stem }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $val = $res.Data.table.Rows[0].sheet_package_id
                if ($val -isnot [DBNull]) { return [guid]$val }
            }
        } catch { }
    }
    if ($folder -and $DocumentName) {
        $resolved = Resolve-SheetPackageFromDocument -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $folder
        if ($resolved.isSheetPackageMember) {
            $ensure = Ensure-SheetPackage -Config $Config -FolderPath $folder -SheetStem $resolved.sheetStem `
                -DocumentRole $resolved.documentRole -DocumentGuid $DocumentGuid -DocumentName $DocumentName
            if ($ensure.IsSuccess -and $ensure.Data -and $ensure.Data.sheetPackageId) {
                try { return [guid]$ensure.Data.sheetPackageId } catch { }
            }
        }
    }
    return $null
}

function Ensure-SheetPackage {
    <#
    .SYNOPSIS
    Upserts one sheet_packages row by (folder_path, sheet_stem) and returns sheet_package_id.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SheetStem,
        [Parameter(Mandatory)][string]$DocumentRole,
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$DesignerEmail,
        [string]$ReviewerEmail,
        [string]$CheckerEmail,
        [string]$QcReviewType,
        [string]$QcAssignedTo,
        [string]$PwStateName,
        [string]$QcCycleId,
        [string]$QcCycleNumber
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'SHEET_PACKAGE_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }
    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath
    $stem = ([string]$SheetStem).Trim()
    $role = ([string]$DocumentRole).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($stem)) {
        return New-QCFailureResult -Code 'SHEET_PACKAGE_IDENTITY_MISSING' -Message 'folder_path and sheet_stem are required.' -Data @{}
    }
    if ($role -notin @('dgn', 'sheet_pdf', 'qc_pdf')) {
        return New-QCSuccessResult -Code 'SHEET_PACKAGE_SKIPPED' -Message 'Document role is not a sheet package member.' -Data @{ written = $false; documentRole = $role }
    }
    $parsedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    $qcLane = ''
    if ($role -eq 'qc_pdf' -and $DocumentName) {
        $qcLane = _QDB-GetQcPdfLaneFromDocumentName -DocumentName $DocumentName
    }
    try {
        $sql = @"
DECLARE @ids TABLE (sheet_package_id UNIQUEIDENTIFIER);
MERGE sheet_packages AS tgt
USING (SELECT @folderPath AS folder_path, @sheetStem AS sheet_stem) AS src
    ON tgt.folder_path = src.folder_path AND tgt.sheet_stem = src.sheet_stem
WHEN MATCHED THEN UPDATE SET
    dgn_guid = CASE WHEN @documentRole = 'dgn' AND @docGuid IS NOT NULL THEN @docGuid ELSE tgt.dgn_guid END,
    dgn_name = CASE WHEN @documentRole = 'dgn' AND @docName IS NOT NULL THEN @docName ELSE tgt.dgn_name END,
    sheet_pdf_guid = CASE WHEN @documentRole = 'sheet_pdf' AND @docGuid IS NOT NULL THEN @docGuid ELSE tgt.sheet_pdf_guid END,
    sheet_pdf_name = CASE WHEN @documentRole = 'sheet_pdf' AND @docName IS NOT NULL THEN @docName ELSE tgt.sheet_pdf_name END,
    qc_pdf_guid = CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'production' AND @docGuid IS NOT NULL THEN @docGuid ELSE tgt.qc_pdf_guid END,
    qc_pdf_name = CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'production' AND @docName IS NOT NULL THEN @docName ELSE tgt.qc_pdf_name END,
    qc_chk_pdf_guid = CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'check' AND @docGuid IS NOT NULL THEN @docGuid ELSE tgt.qc_chk_pdf_guid END,
    qc_chk_pdf_name = CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'check' AND @docName IS NOT NULL THEN @docName ELSE tgt.qc_chk_pdf_name END,
    qc_rev_pdf_guid = CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'review' AND @docGuid IS NOT NULL THEN @docGuid ELSE tgt.qc_rev_pdf_guid END,
    qc_rev_pdf_name = CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'review' AND @docName IS NOT NULL THEN @docName ELSE tgt.qc_rev_pdf_name END,
    designer_email = CASE WHEN @documentRole = 'dgn' THEN COALESCE(@designerEmail, tgt.designer_email)
                          WHEN @documentRole = 'sheet_pdf' THEN COALESCE(@designerEmail, tgt.designer_email)
                          ELSE tgt.designer_email END,
    reviewer_email = CASE WHEN @documentRole = 'dgn' THEN COALESCE(@reviewerEmail, tgt.reviewer_email)
                          WHEN @documentRole = 'sheet_pdf' THEN COALESCE(@reviewerEmail, tgt.reviewer_email)
                          ELSE tgt.reviewer_email END,
    checker_email = CASE WHEN @documentRole = 'dgn' THEN COALESCE(@checkerEmail, tgt.checker_email)
                         WHEN @documentRole = 'sheet_pdf' THEN COALESCE(@checkerEmail, tgt.checker_email)
                         ELSE tgt.checker_email END,
    qc_review_type = CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN COALESCE(@qcReviewType, tgt.qc_review_type) ELSE tgt.qc_review_type END,
    qc_assigned_to = CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN COALESCE(@qcAssignedTo, tgt.qc_assigned_to) ELSE tgt.qc_assigned_to END,
    pw_state_name = tgt.pw_state_name,
    qc_cycle_id = CASE WHEN @documentRole = 'sheet_pdf' THEN COALESCE(@qcCycleId, tgt.qc_cycle_id) ELSE tgt.qc_cycle_id END,
    qc_cycle_number = CASE WHEN @documentRole = 'sheet_pdf' THEN COALESCE(@qcCycleNumber, tgt.qc_cycle_number) ELSE tgt.qc_cycle_number END,
    last_updated_at = SYSDATETIMEOFFSET()
WHEN NOT MATCHED THEN INSERT (
    sheet_package_id, sheet_stem, folder_path,
    dgn_guid, dgn_name, sheet_pdf_guid, sheet_pdf_name, qc_pdf_guid, qc_pdf_name,
    qc_chk_pdf_guid, qc_chk_pdf_name, qc_rev_pdf_guid, qc_rev_pdf_name,
    designer_email, reviewer_email, checker_email, qc_review_type, qc_assigned_to, pw_state_name,
    qc_cycle_id, qc_cycle_number
) VALUES (
    NEWID(), @sheetStem, @folderPath,
    CASE WHEN @documentRole = 'dgn' THEN @docGuid END,
    CASE WHEN @documentRole = 'dgn' THEN @docName END,
    CASE WHEN @documentRole = 'sheet_pdf' THEN @docGuid END,
    CASE WHEN @documentRole = 'sheet_pdf' THEN @docName END,
    CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'production' THEN @docGuid END,
    CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'production' THEN @docName END,
    CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'check' THEN @docGuid END,
    CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'check' THEN @docName END,
    CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'review' THEN @docGuid END,
    CASE WHEN @documentRole = 'qc_pdf' AND @qcLane = 'review' THEN @docName END,
    CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN @designerEmail END,
    CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN @reviewerEmail END,
    CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN @checkerEmail END,
    CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN @qcReviewType END,
    CASE WHEN @documentRole IN ('dgn', 'sheet_pdf') THEN @qcAssignedTo END,
    NULL,
    CASE WHEN @documentRole = 'sheet_pdf' THEN @qcCycleId END,
    CASE WHEN @documentRole = 'sheet_pdf' THEN @qcCycleNumber END
)
OUTPUT inserted.sheet_package_id INTO @ids;
SELECT TOP 1 sheet_package_id FROM @ids;
"@
        $params = @{
            folderPath = $folder
            sheetStem = $stem
            documentRole = $role
            qcLane = if ($qcLane) { $qcLane } else { [DBNull]::Value }
            docGuid = if ($parsedGuid) { $parsedGuid } else { [DBNull]::Value }
            docName = if ($DocumentName) { [string]$DocumentName } else { [DBNull]::Value }
            designerEmail = if ($DesignerEmail) { $DesignerEmail } else { [DBNull]::Value }
            reviewerEmail = if ($ReviewerEmail) { $ReviewerEmail } else { [DBNull]::Value }
            checkerEmail = if ($CheckerEmail) { $CheckerEmail } else { [DBNull]::Value }
            qcReviewType = if ($QcReviewType) { $QcReviewType } else { [DBNull]::Value }
            qcAssignedTo = if ($QcAssignedTo) { $QcAssignedTo } else { [DBNull]::Value }
            pwStateName = if ($PwStateName) { $PwStateName } else { [DBNull]::Value }
            qcCycleId = if ($QcCycleId) { $QcCycleId } else { [DBNull]::Value }
            qcCycleNumber = if ($QcCycleNumber) { $QcCycleNumber } else { [DBNull]::Value }
        }
        $res = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
        $scalarValue = if ($res.Data -and $res.Data.ContainsKey('value')) { $res.Data.value } else { $null }
        if (-not $res.IsSuccess -or $null -eq $scalarValue -or $scalarValue -is [DBNull]) {
            return New-QCErrorResult -Code 'SHEET_PACKAGE_WRITE_FAILED' -Message $res.Message -Data @{ operation = 'ensure_sheet_package' }
        }
        $packageId = [guid][string]$scalarValue
        return New-QCSuccessResult -Code 'SHEET_PACKAGE_WRITTEN' -Message 'Sheet package upserted.' -Data @{
            written = $true
            sheetPackageId = $packageId
            sheetStem = $stem
            documentRole = $role
        }
    } catch {
        return New-QCErrorResult -Code 'SHEET_PACKAGE_EXCEPTION' -Message $_.Exception.Message -Data @{ operation = 'ensure_sheet_package' }
    }
}

function Write-SheetDocument {
    <#
    .SYNOPSIS
    Upserts one sheet_documents row for a physical ProjectWise document.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][guid]$SheetPackageId,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$DocumentRole,
        [string]$PwStateName,
        [string]$Extension,
        [string]$SourceType,
        [string]$FileModifiedAt
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'SHEET_DOCUMENT_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }
    $parsedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    if (-not $parsedGuid) {
        return New-QCFailureResult -Code 'SHEET_DOCUMENT_GUID_INVALID' -Message 'document_guid is not a valid GUID.' -Data @{ documentGuid = $DocumentGuid }
    }
    $role = ([string]$DocumentRole).Trim().ToLowerInvariant()
    if ($role -notin @('dgn', 'sheet_pdf', 'qc_pdf')) {
        return New-QCFailureResult -Code 'SHEET_DOCUMENT_ROLE_INVALID' -Message 'document_role must be dgn, sheet_pdf, or qc_pdf.' -Data @{ documentRole = $role }
    }
    try {
        # Upsert by (sheet_package_id, document_role) so a recreated lane PDF with a new
        # ProjectWise GUID replaces the prior member row instead of failing the role unique key.
        $sql = @"
MERGE sheet_documents AS tgt
USING (SELECT @sheetPackageId AS sheet_package_id, @documentRole AS document_role) AS src
    ON tgt.sheet_package_id = src.sheet_package_id AND tgt.document_role = src.document_role
WHEN MATCHED THEN UPDATE SET
    document_guid = @docGuid,
    document_name = @docName,
    pw_state_name = COALESCE(@pwStateName, tgt.pw_state_name),
    extension = COALESCE(@extension, tgt.extension),
    source_type = COALESCE(@sourceType, tgt.source_type),
    last_seen_at = SYSDATETIMEOFFSET(),
    file_modified_at = COALESCE(@fileModifiedAt, tgt.file_modified_at)
WHEN NOT MATCHED THEN INSERT
    (document_guid, sheet_package_id, document_name, document_role, pw_state_name, extension, source_type, last_seen_at, file_modified_at)
VALUES
    (@docGuid, @sheetPackageId, @docName, @documentRole, @pwStateName, @extension, @sourceType, SYSDATETIMEOFFSET(), @fileModifiedAt);
"@
        $params = @{
            docGuid = $parsedGuid
            sheetPackageId = $SheetPackageId
            docName = [string]$DocumentName
            documentRole = $role
            pwStateName = if ($PwStateName) { $PwStateName } else { [DBNull]::Value }
            extension = if ($Extension) { $Extension } else { [DBNull]::Value }
            sourceType = if ($SourceType) { $SourceType } else { [DBNull]::Value }
            fileModifiedAt = if ($FileModifiedAt) { $FileModifiedAt } else { [DBNull]::Value }
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            return New-QCErrorResult -Code 'SHEET_DOCUMENT_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation = 'write_sheet_document' }
        }
        return New-QCSuccessResult -Code 'SHEET_DOCUMENT_WRITTEN' -Message 'Sheet document upserted.' -Data @{
            written = $true
            sheetPackageId = $SheetPackageId
            documentGuid = $parsedGuid
            documentRole = $role
        }
    } catch {
        return New-QCErrorResult -Code 'SHEET_DOCUMENT_EXCEPTION' -Message $_.Exception.Message -Data @{ operation = 'write_sheet_document' }
    }
}

function _QDB-LinkSheetIndexPackageId {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid,
        [guid]$SheetPackageId
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return }
    $g = ([string]$DocumentGuid).Trim()
    if ($g.Length -lt 5) { return }
    try {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index
SET sheet_package_id = @sheetPackageId,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@ -Parameters @{
            docGuid = $g
            sheetPackageId = $SheetPackageId
        })
    } catch { }
}

function _QDB-SyncSheetPackageDualWrite {
    <#
    .SYNOPSIS
    Phase 1 dual-write: upsert sheet_packages/sheet_documents and link sheet_index.sheet_package_id.
    #>
    param(
        [hashtable]$Config,
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$Extension = '',
        [string]$SourceType = '',
        [string]$DesignerEmail = '',
        [string]$ReviewerEmail = '',
        [string]$CheckerEmail = '',
        [string]$QcReviewType = '',
        [string]$QcAssignedTo = '',
        [string]$PwStateName = '',
        [string]$FileModifiedAt = '',
        [string]$QcCycleId = '',
        [string]$QcCycleNumber = ''
    )
    $resolved = Resolve-SheetPackageFromDocument -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath
    if (-not $resolved.isSheetPackageMember) { return $null }
    $pkgRes = Ensure-SheetPackage -Config $Config `
        -FolderPath $resolved.folderPath `
        -SheetStem $resolved.sheetStem `
        -DocumentRole $resolved.documentRole `
        -DocumentGuid $resolved.documentGuid `
        -DocumentName $resolved.documentName `
        -DesignerEmail $DesignerEmail `
        -ReviewerEmail $ReviewerEmail `
        -CheckerEmail $CheckerEmail `
        -QcReviewType $QcReviewType `
        -QcAssignedTo $QcAssignedTo `
        -PwStateName $PwStateName `
        -QcCycleId $QcCycleId `
        -QcCycleNumber $QcCycleNumber
    if (-not $pkgRes.IsSuccess) { return $null }
    $packageId = $pkgRes.Data.sheetPackageId
    if (-not $packageId) { return $null }
    [void](Write-SheetDocument -Config $Config -SheetPackageId $packageId `
        -DocumentGuid $resolved.documentGuid -DocumentName $resolved.documentName `
        -DocumentRole $resolved.documentRole -PwStateName $PwStateName `
        -Extension $Extension -SourceType $SourceType -FileModifiedAt $FileModifiedAt)
    _QDB-LinkSheetIndexPackageId -Config $Config -DocumentGuid $resolved.documentGuid -SheetPackageId $packageId
    return $packageId
}

function _QDB-NewSheetPackageBackfillConflict {
    param(
        [string]$ConflictType,
        [string]$FolderPath = '',
        [string]$SheetStem = '',
        [string]$DocumentRole = '',
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$Details = ''
    )
    return [ordered]@{
        conflictType = $ConflictType
        folderPath = $FolderPath
        sheetStem = $SheetStem
        documentRole = $DocumentRole
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        details = $Details
    }
}

function _QDB-NewDeterministicSheetPackageId {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SheetStem
    )
    $key = ([string]$FolderPath).Trim().ToLowerInvariant() + '|' + ([string]$SheetStem).Trim().ToLowerInvariant()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes('sheetpkg:' + $key)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    $guidBytes = New-Object byte[] 16
    [Array]::Copy($hash, $guidBytes, 16)
    $guidBytes[6] = ($guidBytes[6] -band 0x0F) -bor 0x40
    $guidBytes[8] = ($guidBytes[8] -band 0x3F) -bor 0x80
    return [guid]$guidBytes
}

function Build-SheetPackageBackfillPlan {
    <#
    .SYNOPSIS
    Builds an idempotent sheet package backfill plan from legacy sheet_index-shaped rows.
    Mirrors scripts/sql/backfill-sheet-packages.sql for offline validation and tests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$SheetIndexRows
    )

    $classified = [System.Collections.Generic.List[object]]::new()
    $conflicts = [System.Collections.Generic.List[object]]::new()
    $seenDocumentGuids = @{}

    foreach ($row in @($SheetIndexRows)) {
        $fp = if ($row.folderPath) { [string]$row.folderPath } elseif ($row.folder_path) { [string]$row.folder_path } else { '' }
        $dn = if ($row.documentName) { [string]$row.documentName } elseif ($row.document_name) { [string]$row.document_name } else { '' }
        $dg = if ($row.documentGuid) { [string]$row.documentGuid } elseif ($row.document_guid) { [string]$row.document_guid } else { '' }

        if ([string]::IsNullOrWhiteSpace($fp)) {
            [void]$conflicts.Add((_QDB-NewSheetPackageBackfillConflict -ConflictType 'missing_folder_path' -DocumentGuid $dg -DocumentName $dn))
            continue
        }
        if (-not (Test-QCSheetIndexFolderPath -FolderPath $fp)) { continue }
        if ([string]::IsNullOrWhiteSpace($dn)) {
            [void]$conflicts.Add((_QDB-NewSheetPackageBackfillConflict -ConflictType 'missing_document_name' -FolderPath $fp -DocumentGuid $dg))
            continue
        }

        $parsed = _QDB-TryParseDocumentGuid -DocumentGuid $dg
        if (-not $parsed) {
            [void]$conflicts.Add((_QDB-NewSheetPackageBackfillConflict -ConflictType 'invalid_guid' -FolderPath $fp -DocumentName $dn -DocumentGuid $dg))
            continue
        }

        $resolved = Resolve-SheetPackageFromDocument -DocumentGuid $dg -DocumentName $dn -FolderPath $fp
        if (-not $resolved.isSheetPackageMember) { continue }

        $guidKey = $parsed.ToString().ToLowerInvariant()
        if ($seenDocumentGuids.ContainsKey($guidKey)) { continue }
        $seenDocumentGuids[$guidKey] = $true

        $classified.Add([pscustomobject]@{
            folderPath = $fp.Trim()
            sheetStem = $resolved.sheetStem
            documentRole = $resolved.documentRole
            documentGuid = $parsed
            documentName = $dn.Trim()
            extension = if ($row.extension) { [string]$row.extension } else { $null }
            sourceType = if ($row.sourceType) { [string]$row.sourceType } elseif ($row.source_type) { [string]$row.source_type } else { $null }
            pwStateName = if ($row.pwStateName) { [string]$row.pwStateName } elseif ($row.pw_state_name) { [string]$row.pw_state_name } else { $null }
            designerEmail = if ($row.designerEmail) { [string]$row.designerEmail } elseif ($row.designer_email) { [string]$row.designer_email } else { $null }
            reviewerEmail = if ($row.reviewerEmail) { [string]$row.reviewerEmail } elseif ($row.reviewer_email) { [string]$row.reviewer_email } else { $null }
            checkerEmail = if ($row.checkerEmail) { [string]$row.checkerEmail } elseif ($row.checker_email) { [string]$row.checker_email } else { $null }
            qcReviewType = if ($row.qcReviewType) { [string]$row.qcReviewType } elseif ($row.qc_review_type) { [string]$row.qc_review_type } else { $null }
            qcAssignedTo = if ($row.qcAssignedTo) { [string]$row.qcAssignedTo } elseif ($row.qc_assigned_to) { [string]$row.qc_assigned_to } else { $null }
            qcCycleId = if ($row.qcCycleId) { [string]$row.qcCycleId } elseif ($row.qc_cycle_id) { [string]$row.qc_cycle_id } else { $null }
            qcCycleNumber = if ($row.qcCycleNumber) { [string]$row.qcCycleNumber } elseif ($row.qc_cycle_number) { [string]$row.qc_cycle_number } else { $null }
            fileModifiedAt = if ($row.fileModifiedAt) { [string]$row.fileModifiedAt } elseif ($row.file_modified_at) { [string]$row.file_modified_at } else { $null }
            productionQcCompletedCount = if ($null -ne $row.productionQcCompletedCount) { [int]$row.productionQcCompletedCount } elseif ($null -ne $row.production_qc_completed_count) { [int]$row.production_qc_completed_count } else { 0 }
            productionQcLastCompletedAt = if ($row.productionQcLastCompletedAt) { $row.productionQcLastCompletedAt } elseif ($row.production_qc_last_completed_at) { $row.production_qc_last_completed_at } else { $null }
            peerReviewCompletedCount = if ($null -ne $row.peerReviewCompletedCount) { [int]$row.peerReviewCompletedCount } elseif ($null -ne $row.peer_review_completed_count) { [int]$row.peer_review_completed_count } else { 0 }
            peerReviewLastCompletedAt = if ($row.peerReviewLastCompletedAt) { $row.peerReviewLastCompletedAt } elseif ($row.peer_review_last_completed_at) { $row.peer_review_last_completed_at } else { $null }
            independentCheckCompletedCount = if ($null -ne $row.independentCheckCompletedCount) { [int]$row.independentCheckCompletedCount } elseif ($null -ne $row.independent_check_completed_count) { [int]$row.independent_check_completed_count } else { 0 }
            independentCheckLastCompletedAt = if ($row.independentCheckLastCompletedAt) { $row.independentCheckLastCompletedAt } elseif ($row.independent_check_last_completed_at) { $row.independent_check_last_completed_at } else { $null }
            qcPdfGuid = if ($row.qcPdfGuid) { [string]$row.qcPdfGuid } elseif ($row.qc_pdf_guid) { [string]$row.qc_pdf_guid } else { $null }
            qcPdfName = if ($row.qcPdfName) { [string]$row.qcPdfName } elseif ($row.qc_pdf_name) { [string]$row.qc_pdf_name } else { $null }
            sourceKind = 'sheet_index_row'
        }) | Out-Null
    }

    foreach ($item in @($classified)) {
        $linkedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $item.qcPdfGuid
        if (-not $linkedGuid) { continue }
        if ([string]::IsNullOrWhiteSpace($item.qcPdfName)) { continue }
        $guidKey = $linkedGuid.ToString().ToLowerInvariant()
        if ($seenDocumentGuids.ContainsKey($guidKey)) { continue }
        $seenDocumentGuids[$guidKey] = $true
        $classified.Add([pscustomobject]@{
            folderPath = $item.folderPath
            sheetStem = $item.sheetStem
            documentRole = 'qc_pdf'
            documentGuid = $linkedGuid
            documentName = ([string]$item.qcPdfName).Trim()
            extension = '.pdf'
            sourceType = $null
            pwStateName = $null
            designerEmail = $null
            reviewerEmail = $null
            checkerEmail = $null
            qcReviewType = $null
            qcAssignedTo = $null
            qcCycleId = $null
            qcCycleNumber = $null
            fileModifiedAt = $null
            productionQcCompletedCount = 0
            productionQcLastCompletedAt = $null
            peerReviewCompletedCount = 0
            peerReviewLastCompletedAt = $null
            independentCheckCompletedCount = 0
            independentCheckLastCompletedAt = $null
            qcPdfGuid = $null
            qcPdfName = $null
            sourceKind = 'linked_qc_pdf'
        }) | Out-Null
    }

    $packageGroups = @($classified | Group-Object { $_.folderPath.ToLowerInvariant() + '|' + $_.sheetStem.ToLowerInvariant() })
    $duplicateRoleKeys = @{}
    foreach ($group in $packageGroups) {
        foreach ($roleGroup in @($group.Group | Group-Object documentRole)) {
            if ($roleGroup.Count -gt 1) {
                $roleKey = $group.Name + '|' + $roleGroup.Name
                $duplicateRoleKeys[$roleKey] = $true
                foreach ($member in @($roleGroup.Group)) {
                    [void]$conflicts.Add((_QDB-NewSheetPackageBackfillConflict `
                        -ConflictType 'duplicate_role' `
                        -FolderPath $member.folderPath `
                        -SheetStem $member.sheetStem `
                        -DocumentRole $member.documentRole `
                        -DocumentGuid $member.documentGuid.ToString() `
                        -DocumentName $member.documentName `
                        -Details ("{0} documents compete for role {1}" -f $roleGroup.Count, $roleGroup.Name)))
                }
            }
        }
    }

    $packages = [System.Collections.Generic.List[object]]::new()
    $documents = [System.Collections.Generic.List[object]]::new()
    $indexLinks = @{}

    foreach ($group in $packageGroups) {
        $sample = $group.Group[0]
        $packageId = _QDB-NewDeterministicSheetPackageId -FolderPath $sample.folderPath -SheetStem $sample.sheetStem
        $pkgKey = $group.Name

        $pickRole = {
            param([string]$Role)
            $roleKey = $pkgKey + '|' + $Role
            if ($duplicateRoleKeys.ContainsKey($roleKey)) { return $null }
            $members = @($group.Group | Where-Object { $_.documentRole -eq $Role })
            if ($members.Count -eq 1) { return $members[0] }
            return $null
        }

        $dgn = & $pickRole 'dgn'
        $pdf = & $pickRole 'sheet_pdf'
        $qc = & $pickRole 'qc_pdf'

        $pwStateName = $null
        if ($qc) { $pwStateName = $qc.pwStateName }
        elseif (-not $qc) {
            if ($pdf) { $pwStateName = $pdf.pwStateName }
            elseif ($dgn) { $pwStateName = $dgn.pwStateName }
        }

        $packages.Add([pscustomobject]@{
            sheetPackageId = $packageId
            folderPath = $sample.folderPath
            sheetStem = $sample.sheetStem
            dgnGuid = if ($dgn) { $dgn.documentGuid } else { $null }
            dgnName = if ($dgn) { $dgn.documentName } else { $null }
            sheetPdfGuid = if ($pdf) { $pdf.documentGuid } else { $null }
            sheetPdfName = if ($pdf) { $pdf.documentName } else { $null }
            qcPdfGuid = if ($qc) { $qc.documentGuid } else { $null }
            qcPdfName = if ($qc) { $qc.documentName } else { $null }
            pwStateName = $pwStateName
            designerEmail = if ($dgn) { $dgn.designerEmail } elseif ($pdf) { $pdf.designerEmail } else { $null }
            reviewerEmail = if ($dgn) { $dgn.reviewerEmail } elseif ($pdf) { $pdf.reviewerEmail } else { $null }
            checkerEmail = if ($dgn) { $dgn.checkerEmail } elseif ($pdf) { $pdf.checkerEmail } else { $null }
            qcReviewType = if ($pdf) { $pdf.qcReviewType } elseif ($dgn) { $dgn.qcReviewType } else { $null }
            qcAssignedTo = if ($pdf) { $pdf.qcAssignedTo } elseif ($dgn) { $dgn.qcAssignedTo } else { $null }
            qcCycleId = if ($pdf) { $pdf.qcCycleId } else { $null }
            qcCycleNumber = if ($pdf) { $pdf.qcCycleNumber } else { $null }
            productionQcCompletedCount = if ($dgn) { $dgn.productionQcCompletedCount } else { 0 }
            productionQcLastCompletedAt = if ($dgn) { $dgn.productionQcLastCompletedAt } else { $null }
            peerReviewCompletedCount = if ($dgn) { $dgn.peerReviewCompletedCount } else { 0 }
            peerReviewLastCompletedAt = if ($dgn) { $dgn.peerReviewLastCompletedAt } else { $null }
            independentCheckCompletedCount = if ($dgn) { $dgn.independentCheckCompletedCount } else { 0 }
            independentCheckLastCompletedAt = if ($dgn) { $dgn.independentCheckLastCompletedAt } else { $null }
        }) | Out-Null

        foreach ($member in @($group.Group)) {
            $roleKey = $pkgKey + '|' + $member.documentRole
            if ($duplicateRoleKeys.ContainsKey($roleKey)) { continue }
            $documents.Add([pscustomobject]@{
                sheetPackageId = $packageId
                documentGuid = $member.documentGuid
                documentName = $member.documentName
                documentRole = $member.documentRole
                pwStateName = $member.pwStateName
                extension = $member.extension
                sourceType = $member.sourceType
                fileModifiedAt = $member.fileModifiedAt
            }) | Out-Null
            $indexLinks[$member.documentGuid.ToString().ToLowerInvariant()] = $packageId
        }
    }

    return [pscustomobject]@{
        packages = @($packages)
        documents = @($documents)
        conflicts = @($conflicts)
        indexLinks = $indexLinks
        packageCount = $packages.Count
        documentCount = $documents.Count
        conflictCount = $conflicts.Count
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
        $packageId = _QDB-SyncSheetPackageDualWrite -Config $Config `
            -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath `
            -Extension $Extension -SourceType $SourceType `
            -DesignerEmail $DesignerEmail -ReviewerEmail $ReviewerEmail -CheckerEmail $CheckerEmail `
            -QcReviewType $QcReviewType -QcAssignedTo $QcAssignedTo -PwStateName $PwStateName -FileModifiedAt $FileModifiedAt
        return New-QCSuccessResult -Code 'SHEET_INDEX_WRITTEN' -Message 'Sheet index upserted.' -Data @{
            written = $true
            rowsAffected = $dbRes.Data.rowsAffected
            sheetPackageId = if ($packageId) { $packageId } else { $null }
        }
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
    if (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try { $PwStateName = Format-QCWorkflowStateName -StateName $PwStateName -Config $Config } catch { }
    }
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
        try {
            $rowRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_name, folder_path, extension, source_type
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
            if ($rowRes.IsSuccess -and $rowRes.Data.table -and $rowRes.Data.table.Rows.Count -gt 0) {
                $row = $rowRes.Data.table.Rows[0]
                $dn = if ($row.document_name -is [DBNull]) { '' } else { [string]$row.document_name }
                $fp = if ($row.folder_path -is [DBNull]) { '' } else { [string]$row.folder_path }
                $ext = if ($row.extension -is [DBNull]) { '' } else { [string]$row.extension }
                $st = if ($row.source_type -is [DBNull]) { '' } else { [string]$row.source_type }
                [void](_QDB-SyncSheetPackageDualWrite -Config $Config -DocumentGuid $DocumentGuid -DocumentName $dn -FolderPath $fp -Extension $ext -SourceType $st -PwStateName $PwStateName)
            }
        } catch { }
        return New-QCSuccessResult -Code 'SHEET_INDEX_WRITTEN' -Message 'Sheet index state updated.' -Data @{ written = $true; rowsAffected = $dbRes.Data.rowsAffected }
    } catch {
        $msg=[string]$_.Exception.Message
        Write-QCJsonLog -Flush -Level 'Error' -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='update_sheet_index_pw_state'; documentGuid=$DocumentGuid }
        return New-QCErrorResult -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation='update_sheet_index_pw_state'; documentGuid=$DocumentGuid }
    }
}

function Get-QCSheetIndexCycle {
    <#
    Returns the active QC cycle for a sheet package (stored on the sheet PDF row).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$SheetStem = ''
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return $null
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace($DocumentGuid)) {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 qc_cycle_id, qc_cycle_number
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = [string]$DocumentGuid.Trim() }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $row = $res.Data.table.Rows[0]
                $cycleId = if ($row.qc_cycle_id -is [DBNull]) { '' } else { [string]$row.qc_cycle_id }
                $cycleNumber = $null
                if ($row.qc_cycle_number -isnot [DBNull]) {
                    $cycleNumber = [string]$row.qc_cycle_number
                    if (-not [string]::IsNullOrWhiteSpace($cycleNumber)) { $cycleNumber = $cycleNumber.Trim() } else { $cycleNumber = $null }
                }
                if (-not [string]::IsNullOrWhiteSpace($cycleId) -or -not [string]::IsNullOrWhiteSpace($cycleNumber)) {
                    return @{ cycleId = $cycleId.Trim(); cycleNumber = $cycleNumber }
                }
            }
        }

        $stem = ([string]$SheetStem).Trim()
        $folder = ([string]$FolderPath).Trim()
        if ($stem.Length -gt 0 -and $folder.Length -gt 0) {
            $pdfName = if ($stem -match '(?i)\.pdf$') { $stem } else { $stem + '.pdf' }
            $res2 = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 qc_cycle_id, qc_cycle_number
FROM sheet_index
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@pdfName)
"@ -Parameters @{ folderPath = $folder; pdfName = $pdfName }
            if ($res2.IsSuccess -and $res2.Data.table -and $res2.Data.table.Rows.Count -gt 0) {
                $row2 = $res2.Data.table.Rows[0]
                $cycleId2 = if ($row2.qc_cycle_id -is [DBNull]) { '' } else { [string]$row2.qc_cycle_id }
                $cycleNumber2 = $null
                if ($row2.qc_cycle_number -isnot [DBNull]) {
                    $cycleNumber2 = [string]$row2.qc_cycle_number
                    if (-not [string]::IsNullOrWhiteSpace($cycleNumber2)) { $cycleNumber2 = $cycleNumber2.Trim() } else { $cycleNumber2 = $null }
                }
                if (-not [string]::IsNullOrWhiteSpace($cycleId2) -or -not [string]::IsNullOrWhiteSpace($cycleNumber2)) {
                    return @{ cycleId = $cycleId2.Trim(); cycleNumber = $cycleNumber2 }
                }
            }
        }
    } catch { }
    return $null
}

function Update-QCSheetIndexCycle {
    <#
    Persists QC cycle identity on the sheet PDF row for notification dedupe and reporting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$CycleId,
        [Parameter(Mandatory)][string]$CycleNumber,
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$SheetStem = ''
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'SHEET_INDEX_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }
    if ([string]::IsNullOrWhiteSpace($CycleId)) {
        return New-QCFailureResult -Code 'SHEET_INDEX_CYCLE_ID_MISSING' -Message 'CycleId is required.' -Data @{}
    }
    $cycleNumberText = [string]$CycleNumber
    if ([string]::IsNullOrWhiteSpace($cycleNumberText)) {
        return New-QCFailureResult -Code 'SHEET_INDEX_CYCLE_NUMBER_MISSING' -Message 'CycleNumber is required.' -Data @{}
    }

    try {
        $params = @{
            cycleId = [string]$CycleId.Trim()
            cycleNumber = $cycleNumberText.Trim()
        }
        $sql = $null
        if (-not [string]::IsNullOrWhiteSpace($DocumentGuid)) {
            $sql = @"
UPDATE sheet_index
SET qc_cycle_id = @cycleId,
    qc_cycle_number = @cycleNumber,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@
            $params['docGuid'] = [string]$DocumentGuid.Trim()
        } else {
            $stem = ([string]$SheetStem).Trim()
            $folder = ([string]$FolderPath).Trim()
            if ($stem.Length -eq 0 -or $folder.Length -eq 0) {
                return New-QCFailureResult -Code 'SHEET_INDEX_CYCLE_TARGET_MISSING' -Message 'DocumentGuid or folderPath+sheetStem is required.' -Data @{}
            }
            $pdfName = if ($stem -match '(?i)\.pdf$') { $stem } else { $stem + '.pdf' }
            $sql = @"
UPDATE sheet_index
SET qc_cycle_id = @cycleId,
    qc_cycle_number = @cycleNumber,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@pdfName)
"@
            $params['folderPath'] = $folder
            $params['pdfName'] = $pdfName
        }

        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            return New-QCErrorResult -Code 'SHEET_INDEX_WRITE_FAILED' -Message $dbRes.Message -Data @{ operation = 'update_sheet_index_cycle' }
        }
        try {
            $lookupSql = if ($params.ContainsKey('docGuid')) {
                "SELECT TOP 1 document_guid, document_name, folder_path, extension, source_type FROM sheet_index WHERE document_guid = @docGuid"
            } else {
                "SELECT TOP 1 document_guid, document_name, folder_path, extension, source_type FROM sheet_index WHERE folder_path = @folderPath AND LOWER(document_name) = LOWER(@pdfName)"
            }
            $rowRes = Invoke-QCDatabaseQuery -Config $Config -Sql $lookupSql -Parameters $params
            if ($rowRes.IsSuccess -and $rowRes.Data.table -and $rowRes.Data.table.Rows.Count -gt 0) {
                $row = $rowRes.Data.table.Rows[0]
                [void](_QDB-SyncSheetPackageDualWrite -Config $Config `
                    -DocumentGuid ([string]$row.document_guid) `
                    -DocumentName ([string]$row.document_name) `
                    -FolderPath ([string]$row.folder_path) `
                    -Extension $(if ($row.extension -is [DBNull]) { '' } else { [string]$row.extension }) `
                    -SourceType $(if ($row.source_type -is [DBNull]) { '' } else { [string]$row.source_type }) `
                    -QcCycleId $params.cycleId -QcCycleNumber $params.cycleNumber)
            }
        } catch { }
        return New-QCSuccessResult -Code 'SHEET_INDEX_WRITTEN' -Message 'Sheet index cycle updated.' -Data @{
            written = $true
            rowsAffected = $dbRes.Data.rowsAffected
            cycleId = $params.cycleId
            cycleNumber = $params.cycleNumber
        }
    } catch {
        $msg = [string]$_.Exception.Message
        return New-QCErrorResult -Code 'SHEET_INDEX_EXCEPTION' -Message $msg -Data @{ operation = 'update_sheet_index_cycle' }
    }
}

function _QDB-GetPwSearchFolderCandidates {
    param([string]$FolderPath)

    $candidates = New-Object System.Collections.ArrayList
    function Add-Candidate([string]$Candidate) {
        if ([string]::IsNullOrWhiteSpace($Candidate)) { return }
        $trimmed = $Candidate.Trim().TrimEnd('\')
        if ($candidates -notcontains $trimmed) { [void]$candidates.Add($trimmed) }
    }

    Add-Candidate $FolderPath
    if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
        Add-Candidate (ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath)
    }
    if (Get-Command -Name 'ConvertTo-PWCanonicalDocumentsFolderPath' -ErrorAction SilentlyContinue) {
        Add-Candidate (ConvertTo-PWCanonicalDocumentsFolderPath -FolderPathProperty $FolderPath)
    }
    if (Get-Command -Name '_PWD-GetSheetRoleFolderCandidates' -ErrorAction SilentlyContinue) {
        foreach ($fp in @(_PWD-GetSheetRoleFolderCandidates -FolderPath $FolderPath)) {
            Add-Candidate $fp
            if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
                Add-Candidate (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp)
            }
        }
    }
    return @($candidates.ToArray())
}

function _QDB-SearchPwLaneQcPdfGuid {
    param(
        [string]$FolderPath,
        [string]$QcPdfName
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or [string]::IsNullOrWhiteSpace($QcPdfName)) { return '' }
    if (-not (Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue)) { return '' }

    $bestGuid = ''
    $bestTicks = [long]::MinValue
    foreach ($searchFolder in @(_QDB-GetPwSearchFolderCandidates -FolderPath $FolderPath)) {
        if ([string]::IsNullOrWhiteSpace($searchFolder)) { continue }
        try {
            $docs = @(Get-PWDocumentsBySearch -FolderPath $searchFolder -DocumentName $QcPdfName -JustThisFolder -ErrorAction SilentlyContinue)
            foreach ($doc in $docs) {
                $g = ''
                try { $g = [string]$doc.DocumentGUID } catch { }
                if ([string]::IsNullOrWhiteSpace($g)) { continue }
                $ticks = [long]::MinValue
                foreach ($n in @('FileUpdatedDate', 'FileUpdateDate', 'DocumentUpdateDate', 'VersionModifiedDate')) {
                    $raw = $null
                    try { if ($doc.PSObject.Properties[$n]) { $raw = $doc.$n } } catch { }
                    if ($null -eq $raw) { continue }
                    try {
                        $dt = [datetime]$raw
                        if ($dt.Ticks -gt $ticks) { $ticks = $dt.Ticks }
                    } catch { }
                }
                if ([string]::IsNullOrWhiteSpace($bestGuid) -or $ticks -gt $bestTicks) {
                    $bestGuid = $g.Trim()
                    $bestTicks = $ticks
                }
            }
        } catch { }
    }
    if (-not [string]::IsNullOrWhiteSpace($bestGuid)) { return $bestGuid }

    if (Get-Command -Name 'Get-PWDocumentsInFolder' -ErrorAction SilentlyContinue) {
        foreach ($searchFolder in @(_QDB-GetPwSearchFolderCandidates -FolderPath $FolderPath)) {
            if ([string]::IsNullOrWhiteSpace($searchFolder)) { continue }
            try {
                $all = @(Get-PWDocumentsInFolder -FolderPath $searchFolder)
                foreach ($doc in $all) {
                    $name = ''
                    try { $name = [string]$doc.Name } catch { }
                    if ([string]::IsNullOrWhiteSpace($name)) { continue }
                    if ($name -ne $QcPdfName -and $name.ToLowerInvariant() -ne $QcPdfName.ToLowerInvariant()) { continue }
                    $g = ''
                    try { $g = [string]$doc.DocumentGUID } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($g)) { return $g.Trim() }
                }
            } catch { }
        }
    }
    return ''
}

function Sync-QCLaneQcPdfGuidFromProjectWise {
    <#
    .SYNOPSIS
    After lane PDF upload/prepend, resolves the live ProjectWise GUID and registers it in telemetry tables.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$QcPdfName,
        [string]$QcProcessType = '',
        [string]$SourceDocumentGuid = '',
        [string]$CurrentPwState = '',
        [switch]$Required
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_LANE_GUID_SYNC_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }

    $qcName = ([string]$QcPdfName).Trim()
    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath
    if ([string]::IsNullOrWhiteSpace($qcName) -or [string]::IsNullOrWhiteSpace($folder)) {
        return New-QCFailureResult -Code 'QC_LANE_GUID_SYNC_INVALID_INPUT' -Message 'Folder path and QC PDF name are required.' -Data @{}
    }

    $processType = ([string]$QcProcessType).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($processType)) {
        $processType = _QDB-GetQcPdfLaneFromDocumentName -DocumentName $qcName
    }
    if ($processType -notin @('production', 'check', 'review')) {
        return New-QCFailureResult -Code 'QC_LANE_GUID_SYNC_INVALID_TYPE' -Message 'qc_process_type must be production, check, or review.' -Data @{ qcPdfName = $qcName }
    }

    $liveGuid = _QDB-SearchPwLaneQcPdfGuid -FolderPath $folder -QcPdfName $qcName
    if ([string]::IsNullOrWhiteSpace($liveGuid)) {
        if ($Required) {
            return New-QCFailureResult -Code 'QC_LANE_GUID_SYNC_PW_NOT_FOUND' -Message 'Live lane QC PDF GUID could not be resolved from ProjectWise.' -Data @{
                folderPath = $folder; qcPdfName = $qcName; qcProcessType = $processType
            }
        }
        return New-QCSuccessResult -Code 'QC_LANE_GUID_SYNC_PW_NOT_FOUND' -Message 'Lane QC PDF not found in ProjectWise yet.' -Data @{
            written = $false; folderPath = $folder; qcPdfName = $qcName; qcProcessType = $processType
        }
    }

    $packageId = $null
    if (-not [string]::IsNullOrWhiteSpace($SourceDocumentGuid)) {
        $packageId = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $SourceDocumentGuid
    }
    if (-not $packageId) {
        $resolved = Resolve-SheetPackageFromDocument -DocumentName $qcName -FolderPath $folder -DocumentGuid $liveGuid
        if ($resolved -and $resolved.sheetPackageId) {
            try { $packageId = [guid]$resolved.sheetPackageId } catch { }
        }
    }
    if (-not $packageId) {
        $pkgRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sheet_package_id
FROM sheet_packages
WHERE folder_path = @folderPath
ORDER BY last_updated_at DESC
"@ -Parameters @{ folderPath = $folder }
        if ($pkgRes.IsSuccess -and $pkgRes.Data.table -and $pkgRes.Data.table.Rows.Count -gt 0) {
            try { $packageId = [guid]$pkgRes.Data.table.Rows[0].sheet_package_id } catch { }
        }
    }
    if (-not $packageId) {
        return New-QCFailureResult -Code 'QC_LANE_GUID_SYNC_NO_PACKAGE' -Message 'Sheet package could not be resolved for lane QC PDF registration.' -Data @{
            folderPath = $folder; qcPdfName = $qcName; documentGuid = $liveGuid
        }
    }

    $upsert = Upsert-SheetPackageQcPdf -Config $Config -SheetPackageId $packageId -QcProcessType $processType `
        -DocumentGuid $liveGuid -DocumentName $qcName -FolderPath $folder -CurrentPwState $CurrentPwState -Required
    if (-not $upsert.IsSuccess) { return $upsert }

    $colGuid = switch ($processType) {
        'check' { 'qc_chk_pdf_guid' }
        'review' { 'qc_rev_pdf_guid' }
        default { 'qc_pdf_guid' }
    }
    $colName = switch ($processType) {
        'check' { 'qc_chk_pdf_name' }
        'review' { 'qc_rev_pdf_name' }
        default { 'qc_pdf_name' }
    }
    $parsedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $liveGuid
    if ($parsedGuid) {
        [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_packages
SET $colGuid = @docGuid, $colName = @docName, last_updated_at = SYSDATETIMEOFFSET()
WHERE sheet_package_id = @sheetPackageId
"@ -Parameters @{
            docGuid = $parsedGuid
            docName = $qcName
            sheetPackageId = $packageId
        })
        [void](Write-SheetDocument -Config $Config -SheetPackageId $packageId `
            -DocumentGuid $liveGuid -DocumentName $qcName -DocumentRole 'qc_pdf')
    }

    if (-not [string]::IsNullOrWhiteSpace($SourceDocumentGuid)) {
        [void](Update-QCSheetQcPdf -Config $Config -SourceDocumentGuid $SourceDocumentGuid `
            -QcPdfGuid $liveGuid -QcPdfName $qcName -QcProcessType $processType)
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_LANE_PDF_GUID_REGISTERED' `
            -Message 'Registered live lane QC PDF GUID from ProjectWise after prepend.' -Data @{
            folderPath = $folder
            qcPdfName = $qcName
            qcProcessType = $processType
            documentGuid = $liveGuid
            sheetPackageId = $packageId.ToString()
            sourceDocumentGuid = $SourceDocumentGuid
        } | Out-Null
    }

    return New-QCSuccessResult -Code 'QC_LANE_PDF_GUID_REGISTERED' -Message 'Live lane QC PDF GUID registered.' -Data @{
        written = $true
        documentGuid = $liveGuid
        documentName = $qcName
        qcProcessType = $processType
        sheetPackageId = $packageId.ToString()
        folderPath = $folder
    }
}

function Resolve-QCSheetQcPdfGuid {
    <#
    .SYNOPSIS
    Resolves the live ProjectWise GUID for a lane QC PDF (*-prod/-chk/-rev.pdf).
    Prefers sheet_package_qc_pdfs, then sheet_index / production alias columns on sheet_packages.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$QcPdfName,
        [string]$SourceDocumentGuid = ''
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return '' }
    $folder = [string]$FolderPath
    $qcName = [string]$QcPdfName
    if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($qcName)) { return '' }
    $lane = _QDB-GetQcPdfLaneFromDocumentName -DocumentName $qcName

    $pwGuid = _QDB-SearchPwLaneQcPdfGuid -FolderPath $folder -QcPdfName $qcName
    if (-not [string]::IsNullOrWhiteSpace($pwGuid)) { return $pwGuid.Trim() }

    try {
        if ($lane) {
            $laneRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_guid
FROM sheet_package_qc_pdfs
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@qcPdfName)
  AND qc_process_type = @qcProcessType
  AND is_active = 1
ORDER BY updated_at DESC
"@ -Parameters @{ folderPath = $folder; qcPdfName = $qcName; qcProcessType = $lane }
            if ($laneRes.IsSuccess -and $laneRes.Data.table -and $laneRes.Data.table.Rows.Count -gt 0) {
                $lg = if ($laneRes.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$laneRes.Data.table.Rows[0].document_guid }
                if (-not [string]::IsNullOrWhiteSpace($lg)) { return $lg.Trim() }
            }
        }
        $idxRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_guid
FROM sheet_index
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@qcPdfName)
ORDER BY
  CASE WHEN sheet_package_id IS NOT NULL THEN 0 ELSE 1 END,
  last_updated_at DESC
"@ -Parameters @{ folderPath = $folder; qcPdfName = $qcName }
        if ($idxRes.IsSuccess -and $idxRes.Data.table -and $idxRes.Data.table.Rows.Count -gt 0) {
            $g = if ($idxRes.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$idxRes.Data.table.Rows[0].document_guid }
            if (-not [string]::IsNullOrWhiteSpace($g)) { return $g.Trim() }
        }
        if (-not [string]::IsNullOrWhiteSpace($SourceDocumentGuid)) {
            $pkgRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sp.qc_pdf_guid
FROM sheet_packages sp
INNER JOIN sheet_index si ON si.sheet_package_id = sp.sheet_package_id
WHERE si.document_guid = @sourceDocGuid
  AND sp.qc_pdf_guid IS NOT NULL
"@ -Parameters @{ sourceDocGuid = [string]$SourceDocumentGuid }
            if ($pkgRes.IsSuccess -and $pkgRes.Data.table -and $pkgRes.Data.table.Rows.Count -gt 0) {
                $g2 = if ($pkgRes.Data.table.Rows[0].qc_pdf_guid -is [DBNull]) { '' } else { [string]$pkgRes.Data.table.Rows[0].qc_pdf_guid }
                if (-not [string]::IsNullOrWhiteSpace($g2)) { return $g2.Trim() }
            }
        }
        $pkgCol = switch ($lane) {
            'check' { 'qc_chk_pdf_guid' }
            'review' { 'qc_rev_pdf_guid' }
            default { 'qc_pdf_guid' }
        }
        $pkgNameCol = switch ($lane) {
            'check' { 'qc_chk_pdf_name' }
            'review' { 'qc_rev_pdf_name' }
            default { 'qc_pdf_name' }
        }
        if (-not $lane) { return '' }
        $pkgNameRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 $pkgCol AS lane_guid
FROM sheet_packages
WHERE folder_path = @folderPath
  AND $pkgCol IS NOT NULL
  AND LOWER($pkgNameCol) = LOWER(@qcPdfName)
ORDER BY last_updated_at DESC
"@ -Parameters @{ folderPath = $folder; qcPdfName = $qcName }
        if ($pkgNameRes.IsSuccess -and $pkgNameRes.Data.table -and $pkgNameRes.Data.table.Rows.Count -gt 0) {
            $g3 = if ($pkgNameRes.Data.table.Rows[0].lane_guid -is [DBNull]) { '' } else { [string]$pkgNameRes.Data.table.Rows[0].lane_guid }
            if (-not [string]::IsNullOrWhiteSpace($g3)) { return $g3.Trim() }
        }
    } catch { }
    return ''
}

function Update-QCSheetQcPdf {
    <#
    .SYNOPSIS
    Links a lane PDF to its source sheet in the sheet_index table. Fire-and-forget.
    Called after a successful QC_PREPEND job.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$SourceDocumentGuid,
        [string]$QcPdfGuid,
        [string]$QcPdfName,
        [string]$QcProcessType = ''
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
        try {
            $packageId = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $SourceDocumentGuid
            if (-not $packageId) {
                $srcRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_name, folder_path, extension, source_type
FROM sheet_index
WHERE document_guid = @sourceDocGuid
"@ -Parameters @{ sourceDocGuid = $SourceDocumentGuid }
                if ($srcRes.IsSuccess -and $srcRes.Data.table -and $srcRes.Data.table.Rows.Count -gt 0) {
                    $srcRow = $srcRes.Data.table.Rows[0]
                    $packageId = _QDB-SyncSheetPackageDualWrite -Config $Config `
                        -DocumentGuid $SourceDocumentGuid `
                        -DocumentName ([string]$srcRow.document_name) `
                        -FolderPath ([string]$srcRow.folder_path) `
                        -Extension $(if ($srcRow.extension -is [DBNull]) { '' } else { [string]$srcRow.extension }) `
                        -SourceType $(if ($srcRow.source_type -is [DBNull]) { '' } else { [string]$srcRow.source_type })
                }
            }
            if ($packageId -and -not [string]::IsNullOrWhiteSpace($QcPdfGuid) -and -not [string]::IsNullOrWhiteSpace($QcPdfName)) {
                $qcParsed = _QDB-TryParseDocumentGuid -DocumentGuid $QcPdfGuid
                if ($qcParsed) {
                    $lane = ([string]$QcProcessType).Trim().ToLowerInvariant()
                    if ([string]::IsNullOrWhiteSpace($lane)) {
                        $lane = _QDB-GetQcPdfLaneFromDocumentName -DocumentName $QcPdfName
                    }
                    $colGuid = switch ($lane) {
                        'check' { 'qc_chk_pdf_guid' }
                        'review' { 'qc_rev_pdf_guid' }
                        default { 'qc_pdf_guid' }
                    }
                    $colName = switch ($lane) {
                        'check' { 'qc_chk_pdf_name' }
                        'review' { 'qc_rev_pdf_name' }
                        default { 'qc_pdf_name' }
                    }
                    [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_packages
SET qc_pdf_guid = CASE WHEN @lane = 'production' THEN @qcPdfGuid ELSE qc_pdf_guid END,
    qc_pdf_name = CASE WHEN @lane = 'production' THEN @qcPdfName ELSE qc_pdf_name END,
    $colGuid = @qcPdfGuid,
    $colName = @qcPdfName,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE sheet_package_id = @sheetPackageId
"@ -Parameters @{
                        qcPdfGuid = $qcParsed
                        qcPdfName = $QcPdfName
                        sheetPackageId = $packageId
                        lane = if ($lane) { $lane } else { 'production' }
                    })
                    [void](Write-SheetDocument -Config $Config -SheetPackageId $packageId `
                        -DocumentGuid $QcPdfGuid -DocumentName $QcPdfName -DocumentRole 'qc_pdf')
                    if ($lane -and (Get-Command -Name 'Upsert-SheetPackageQcPdf' -ErrorAction SilentlyContinue)) {
                        [void](Upsert-SheetPackageQcPdf -Config $Config -SheetPackageId $packageId -QcProcessType $lane `
                            -DocumentGuid $QcPdfGuid -DocumentName $QcPdfName -FolderPath '' -Required)
                    }
                }
            }
        } catch { }
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
            foreach ($row in $Rows) {
                $ext = $null
                if ($row.documentName) {
                    $extPart = [System.IO.Path]::GetExtension([string]$row.documentName)
                    if ($extPart) { $ext = $extPart.ToLowerInvariant() }
                }
                [void](_QDB-SyncSheetPackageDualWrite -Config $Config `
                    -DocumentGuid ([string]$row.documentGuid) `
                    -DocumentName ([string]$row.documentName) `
                    -FolderPath ([string]$row.folderPath) `
                    -Extension $ext `
                    -SourceType $(if ($row.sourceType) { [string]$row.sourceType } else { '' }) `
                    -DesignerEmail $(if ($row.designerEmail) { [string]$row.designerEmail } else { '' }) `
                    -ReviewerEmail $(if ($row.reviewerEmail) { [string]$row.reviewerEmail } else { '' }) `
                    -CheckerEmail $(if ($row.checkerEmail) { [string]$row.checkerEmail } else { '' }) `
                    -QcReviewType $(if ($row.qcReviewType) { [string]$row.qcReviewType } else { '' }) `
                    -QcAssignedTo $(if ($row.qcAssignedTo) { [string]$row.qcAssignedTo } else { '' }) `
                    -PwStateName $(if ($row.pwStateName) { [string]$row.pwStateName } else { '' }))
            }
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

function Get-QCNewerSheetDocumentStateAuditEvent {
    <#
    .SYNOPSIS
    Returns the earliest DOCUMENT_STATE audit_events row newer than the current event for the same sheet group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [array]$MemberDocumentNames = @(),
        [array]$MemberDocumentGuids = @(),
        [Nullable[long]]$CurrentAuditEventId = $null,
        [string]$CurrentAuditEventAt = ''
    )

    $notFound = @{ found = $false }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_SKIPPED' -Message 'Database disabled.' -Data $notFound
    }

    $names = @($MemberDocumentNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    $guids = @($MemberDocumentGuids | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    if ($names.Count -eq 0 -and $guids.Count -eq 0) {
        return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_NONE' -Message 'No sheet member identities supplied.' -Data $notFound
    }

    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath
    if ([string]::IsNullOrWhiteSpace($folder)) { $folder = [string]$FolderPath }

    $memberClauses = [System.Collections.Generic.List[string]]::new()
    $params = @{ folderPath = $folder }
    $ni = 0
    foreach ($n in $names) {
        $key = "docName$ni"
        $params[$key] = $n
        [void]$memberClauses.Add("LOWER(ae.pw_itemname) = LOWER(@$key)")
        $ni++
    }
    $gi = 0
    foreach ($g in $guids) {
        $key = "docGuid$gi"
        $params[$key] = $g
        [void]$memberClauses.Add("ae.pw_objguid = @$key")
        $gi++
    }
    if ($memberClauses.Count -eq 0) {
        return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_NONE' -Message 'No member match clauses.' -Data $notFound
    }

    $newerClause = ''
    if ($null -ne $CurrentAuditEventId -and $CurrentAuditEventId -gt 0) {
        $newerClause = 'AND ae.id > @currentAuditEventId'
        $params['currentAuditEventId'] = [long]$CurrentAuditEventId
    } elseif (-not [string]::IsNullOrWhiteSpace($CurrentAuditEventAt)) {
        $currentAtUtc = $null
        if (Get-Command -Name '_QCAT-ParsePwActTimeUtc' -ErrorAction SilentlyContinue) {
            $currentAtUtc = _QCAT-ParsePwActTimeUtc -ActTime $CurrentAuditEventAt
        }
        if ($null -eq $currentAtUtc) {
            return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_NO_ANCHOR' -Message 'Current audit timestamp could not be normalized to UTC.' -Data $notFound
        }
        $newerClause = 'AND ae.pw_acttime > @currentAuditEventAt'
        $params['currentAuditEventAt'] = $currentAtUtc.ToString('yyyy-MM-dd HH:mm:ss')
    } else {
        return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_NO_ANCHOR' -Message 'No audit event anchor for recency comparison.' -Data $notFound
    }

    $sql = @"
SELECT TOP 1 ae.id, ae.pw_acttime, ae.pw_objguid, ae.pw_itemname, ae.pw_userno, ae.processed
FROM audit_events ae
WHERE ae.pw_action = 1012
  AND ae.resolved_folder = @folderPath
  AND ($($memberClauses -join ' OR '))
  $newerClause
ORDER BY ae.id ASC
"@
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $res.IsSuccess -or -not $res.Data -or -not $res.Data.table -or $res.Data.table.Rows.Count -eq 0) {
            return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_NOT_FOUND' -Message 'No newer DOCUMENT_STATE audit event for sheet group.' -Data $notFound
        }
        $r = $res.Data.table.Rows[0]
        $candidateId = if ($r.id -is [DBNull]) { $null } else { [long]$r.id }
        $candidateTime = if ($r.pw_acttime -is [DBNull]) { '' } else { [string]$r.pw_acttime }
        $candidateName = if ($r.pw_itemname -is [DBNull]) { '' } else { [string]$r.pw_itemname }
        $isStrictlyNewer = $false
        if (Get-Command -Name '_QCAT-TestAuditTimeIsStrictlyAfterUtc' -ErrorAction SilentlyContinue) {
            $isStrictlyNewer = _QCAT-TestAuditTimeIsStrictlyAfterUtc -CandidateTime $candidateTime -CurrentTime $CurrentAuditEventAt `
                -CandidateAuditEventId $candidateId -CurrentAuditEventId $CurrentAuditEventId
        } elseif ($null -ne $CurrentAuditEventId -and $CurrentAuditEventId -gt 0 -and $null -ne $candidateId) {
            $isStrictlyNewer = ([long]$candidateId -gt [long]$CurrentAuditEventId)
        }
        if (-not $isStrictlyNewer) {
            $rejected = @{
                found = $false
                rejectedBlockingReason = 'blocking_candidate_older_than_current'
                rejectedBlockingAuditEventId = $candidateId
                rejectedBlockingAuditTime = $candidateTime
                rejectedBlockingDocumentName = $candidateName
            }
            return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_REJECTED' -Message 'Candidate DOCUMENT_STATE audit event is not newer than current anchor.' -Data $rejected
        }
        $data = @{
            found = $true
            id = $candidateId
            pwActtime = $candidateTime
            documentGuid = if ($r.pw_objguid -is [DBNull]) { '' } else { [string]$r.pw_objguid }
            documentName = $candidateName
            changedByUser = if ($r.pw_userno -is [DBNull]) { $null } else { [int]$r.pw_userno }
            processed = if ($r.processed -is [DBNull]) { $false } else { [bool]$r.processed }
        }
        return New-QCSuccessResult -Code 'NEWER_STATE_AUDIT_FOUND' -Message 'Newer DOCUMENT_STATE audit event found for sheet group.' -Data $data
    } catch {
        return New-QCFailureResult -Code 'NEWER_STATE_AUDIT_QUERY_FAILED' -Message $_.Exception.Message -Data $notFound
    }
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

function Get-QCReviewTypeBucket {
    <#
    .SYNOPSIS
    Maps QC process/review type labels to normalized buckets: production, check, review.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReviewType
    )

    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        return (Normalize-QCProcessType -ProcessType $ReviewType)
    }

    $norm = ([string]$ReviewType).Trim().ToLowerInvariant()
    switch -Regex ($norm) {
        '^production(\s+qc)?$|^production$|^qc$' { return 'production' }
        '^check$|^independent(\s+(check|review))?$|^independent_check$|^independent$|^ic$' { return 'check' }
        '^review$|^peer(\s+review)?$|^peer_review$|^peer$' { return 'review' }
        default { return $null }
    }
}

function Resolve-QCCycleCompletionSheetPackageId {
    <#
    .SYNOPSIS
    Resolves sheet_package_id for QC cycle completion writes and rollups.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$DocumentGuid = '',
        [Nullable[guid]]$SheetPackageId = $null
    )
    return (_QDB-ResolveQCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $DocumentGuid -SheetPackageId $SheetPackageId)
}

function _QDB-ResolveQCCycleCompletionSheetPackageId {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid = '',
        [Nullable[guid]]$SheetPackageId = $null
    )
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) {
        return $SheetPackageId
    }
    $pkg = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $DocumentGuid
    if ($pkg) { return $pkg }
    $parsed = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    if (-not $parsed) { return $null }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sheet_package_id
FROM sheet_packages
WHERE dgn_guid = @docGuid
"@ -Parameters @{ docGuid = $parsed }
        if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
            $val = $res.Data.table.Rows[0].sheet_package_id
            if ($val -isnot [DBNull]) { return [guid]$val }
        }
    } catch { }
    return $null
}

function _QDB-FindExistingQCCycleCompletion {
    param(
        [hashtable]$Config,
        [Nullable[guid]]$SheetPackageId,
        [string]$DocumentGuid,
        [string]$QcCycleId,
        [string]$QcReviewType
    )
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) {
        $byPackage = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 id, completed_at, sheet_package_id
FROM qc_cycle_completions
WHERE sheet_package_id = @sheetPackageId
  AND qc_cycle_id = @qcCycleId
  AND qc_review_type = @qcReviewType
"@ -Parameters @{
            sheetPackageId = $SheetPackageId
            qcCycleId = $QcCycleId
            qcReviewType = $QcReviewType
        }
        if ($byPackage.IsSuccess -and $byPackage.Data.table -and $byPackage.Data.table.Rows.Count -gt 0) {
            return $byPackage.Data.table.Rows[0]
        }
    }
    $byDoc = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 id, completed_at, sheet_package_id
FROM qc_cycle_completions
WHERE document_guid = @documentGuid
  AND qc_cycle_id = @qcCycleId
  AND qc_review_type = @qcReviewType
"@ -Parameters @{
        documentGuid = $DocumentGuid
        qcCycleId = $QcCycleId
        qcReviewType = $QcReviewType
    }
    if ($byDoc.IsSuccess -and $byDoc.Data.table -and $byDoc.Data.table.Rows.Count -gt 0) {
        return $byDoc.Data.table.Rows[0]
    }
    return $null
}

function Ensure-QCCycleCompletion {
    <#
    .SYNOPSIS
    Idempotently records a completed QC cycle keyed by sheet_package_id + qc_cycle_id + qc_review_type.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$QcCycleId,
        [Parameter(Mandatory)][string]$QcReviewType,
        [Nullable[guid]]$SheetPackageId = $null,
        [string]$DocumentName = '',
        [Nullable[int]]$QcCycleNumber = $null,
        [Nullable[long]]$TransitionEventId = $null,
        [Nullable[long]]$AuditEventId = $null,
        [string]$CompletedBy = '',
        [Nullable[datetime]]$CompletedAt = $null
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_CYCLE_COMPLETION_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ inserted = $false; reused = $false; completionId = $null; sheetPackageId = $null }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_CYCLE_COMPLETION_PLANNED' -Message 'Dry-run: QC cycle completion not written.' -Data @{ inserted = $false; planned = $true; reused = $false; completionId = $null; sheetPackageId = $SheetPackageId }
    }

    $docGuid = ([string]$DocumentGuid).Trim().Trim('{}')
    $cycleId = ([string]$QcCycleId).Trim()
    $reviewType = ([string]$QcReviewType).Trim()
    if ([string]::IsNullOrWhiteSpace($docGuid) -or [string]::IsNullOrWhiteSpace($cycleId) -or [string]::IsNullOrWhiteSpace($reviewType)) {
        return New-QCFailureResult -Code 'QC_CYCLE_COMPLETION_INVALID' -Message 'DocumentGuid, QcCycleId, and QcReviewType are required.' -Data @{ inserted = $false; reused = $false; completionId = $null; sheetPackageId = $null }
    }

    $packageId = _QDB-ResolveQCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $docGuid -SheetPackageId $SheetPackageId
    if (-not $packageId) {
        return New-QCSuccessResult -Code 'QC_CYCLE_COMPLETION_SKIPPED' -Message 'sheet_package_id could not be resolved for completion insert.' -Data @{
            inserted = $false; reused = $false; completionId = $null; sheetPackageId = $null; reason = 'sheet_package_not_found'
        }
    }

    try {
        $existingRow = _QDB-FindExistingQCCycleCompletion -Config $Config -SheetPackageId $packageId `
            -DocumentGuid $docGuid -QcCycleId $cycleId -QcReviewType $reviewType
        if ($existingRow) {
            if ($existingRow.sheet_package_id -is [DBNull] -or $null -eq $existingRow.sheet_package_id) {
                try {
                    [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE qc_cycle_completions
SET sheet_package_id = @sheetPackageId
WHERE id = @id AND sheet_package_id IS NULL
"@ -Parameters @{ sheetPackageId = $packageId; id = [long]$existingRow.id })
                } catch { }
            }
            return New-QCSuccessResult -Code 'QC_CYCLE_COMPLETION_REUSED' -Message 'Existing qc_cycle_completions row reused.' -Data @{
                inserted = $false; reused = $true; completionId = [long]$existingRow.id
                completedAt = if ($existingRow.completed_at -is [DBNull]) { $null } else { [datetime]$existingRow.completed_at }
                sheetPackageId = $packageId
            }
        }
    } catch { }

    $completedAtValue = if ($null -ne $CompletedAt) { $CompletedAt } else { [datetime]::UtcNow }
    $parsedDocGuid = _QDB-TryParseDocumentGuid -DocumentGuid $docGuid
    try {
        $sql = @"
INSERT INTO qc_cycle_completions
    (sheet_package_id, document_guid, document_name, qc_cycle_id, qc_cycle_number, qc_review_type, completed_at, completed_by, transition_event_id, audit_event_id)
OUTPUT INSERTED.id
VALUES
    (@sheetPackageId, @documentGuid, @documentName, @qcCycleId, @qcCycleNumber, @qcReviewType, @completedAt, @completedBy, @transitionEventId, @auditEventId)
"@
        $params = @{
            sheetPackageId = $packageId
            documentGuid = if ($parsedDocGuid) { $parsedDocGuid } else { [guid]$docGuid }
            documentName = if ($DocumentName) { $DocumentName } else { $null }
            qcCycleId = $cycleId
            qcCycleNumber = if ($null -ne $QcCycleNumber) { $QcCycleNumber } else { $null }
            qcReviewType = $reviewType
            completedAt = $completedAtValue
            completedBy = if ($CompletedBy) { $CompletedBy } else { $null }
            transitionEventId = if ($null -ne $TransitionEventId) { $TransitionEventId } else { $null }
            auditEventId = if ($null -ne $AuditEventId) { $AuditEventId } else { $null }
        }
        $dbRes = Invoke-QCDatabaseScalar -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            $msg = [string]$dbRes.Message
            if ($msg -match 'UQ_qc_cycle_completions_package|UQ_qc_cycle_completions_cycle|duplicate key|2601|2627') {
                try {
                    $dupRow = _QDB-FindExistingQCCycleCompletion -Config $Config -SheetPackageId $packageId `
                        -DocumentGuid $docGuid -QcCycleId $cycleId -QcReviewType $reviewType
                    if ($dupRow) {
                        return New-QCSuccessResult -Code 'QC_CYCLE_COMPLETION_REUSED' -Message 'QC cycle completion duplicate suppressed by unique constraint.' -Data @{
                            inserted = $false; reused = $true; completionId = [long]$dupRow.id
                            completedAt = if ($dupRow.completed_at -is [DBNull]) { $null } else { [datetime]$dupRow.completed_at }
                            sheetPackageId = $packageId
                        }
                    }
                } catch { }
            }
            return New-QCErrorResult -Code 'QC_CYCLE_COMPLETION_WRITE_FAILED' -Message $msg -Data @{ inserted = $false; reused = $false; completionId = $null; sheetPackageId = $packageId }
        }
        $completionId = $null
        if ($null -ne $dbRes.Data.value) {
            try { $completionId = [long]$dbRes.Data.value } catch { $completionId = $null }
        }
        return New-QCSuccessResult -Code 'QC_CYCLE_COMPLETION_WRITTEN' -Message 'qc_cycle_completions row inserted.' -Data @{
            inserted = $true; reused = $false; completionId = $completionId; completedAt = $completedAtValue; sheetPackageId = $packageId
        }
    } catch {
        return New-QCErrorResult -Code 'QC_CYCLE_COMPLETION_EXCEPTION' -Message $_.Exception.Message -Data @{ inserted = $false; reused = $false; completionId = $null; sheetPackageId = $packageId }
    }
}

function Update-QCSheetCycleCompletionSummary {
    <#
    .SYNOPSIS
    Rebuilds sheet_packages QC completion summary columns from qc_cycle_completions.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$DocumentGuid = '',
        [Nullable[guid]]$SheetPackageId = $null
    )

    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_CYCLE_SUMMARY_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'QC_CYCLE_SUMMARY_PLANNED' -Message 'Dry-run: sheet_packages completion summary not updated.' -Data @{ written = $false; planned = $true }
    }

    $packageId = _QDB-ResolveQCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $DocumentGuid -SheetPackageId $SheetPackageId
    if (-not $packageId) {
        return New-QCFailureResult -Code 'QC_CYCLE_SUMMARY_INVALID' -Message 'SheetPackageId could not be resolved.' -Data @{ written = $false }
    }

    try {
        $aggRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT qc_review_type, COUNT(*) AS completed_count, MAX(completed_at) AS last_completed_at
FROM qc_cycle_completions
WHERE sheet_package_id = @sheetPackageId
GROUP BY qc_review_type
"@ -Parameters @{ sheetPackageId = $packageId }

        $productionCount = 0; $productionLast = $null
        $peerCount = 0; $peerLast = $null
        $independentCount = 0; $independentLast = $null

        if ($aggRes.IsSuccess -and $aggRes.Data.table) {
            foreach ($row in @($aggRes.Data.table.Rows)) {
                $rt = if ($row.qc_review_type -is [DBNull]) { '' } else { [string]$row.qc_review_type }
                $bucket = Get-QCReviewTypeBucket -ReviewType $rt
                if (-not $bucket) { continue }
                $cnt = 0
                try { $cnt = [int]$row.completed_count } catch { $cnt = 0 }
                $lastAt = $null
                if ($row.last_completed_at -isnot [DBNull]) {
                    try { $lastAt = [datetime]$row.last_completed_at } catch { $lastAt = $null }
                }
                switch ($bucket) {
                    'production' {
                        $productionCount += $cnt
                        if ($null -eq $productionLast -or ($null -ne $lastAt -and $lastAt -gt $productionLast)) { $productionLast = $lastAt }
                    }
                    'review' {
                        $peerCount += $cnt
                        if ($null -eq $peerLast -or ($null -ne $lastAt -and $lastAt -gt $peerLast)) { $peerLast = $lastAt }
                    }
                    'check' {
                        $independentCount += $cnt
                        if ($null -eq $independentLast -or ($null -ne $lastAt -and $lastAt -gt $independentLast)) { $independentLast = $lastAt }
                    }
                }
            }
        }

        $upd = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_packages
SET production_qc_completed_count = @productionCount,
    production_qc_last_completed_at = @productionLast,
    peer_review_completed_count = @peerCount,
    peer_review_last_completed_at = @peerLast,
    independent_check_completed_count = @independentCount,
    independent_check_last_completed_at = @independentLast,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE sheet_package_id = @sheetPackageId
"@ -Parameters @{
            sheetPackageId = $packageId
            productionCount = $productionCount
            productionLast = if ($null -ne $productionLast) { $productionLast } else { $null }
            peerCount = $peerCount
            peerLast = if ($null -ne $peerLast) { $peerLast } else { $null }
            independentCount = $independentCount
            independentLast = if ($null -ne $independentLast) { $independentLast } else { $null }
        }

        if (-not $upd.IsSuccess) {
            return New-QCErrorResult -Code 'QC_CYCLE_SUMMARY_UPDATE_FAILED' -Message $upd.Message -Data @{ written = $false; sheetPackageId = $packageId }
        }
        return New-QCSuccessResult -Code 'QC_CYCLE_SUMMARY_UPDATED' -Message 'sheet_packages completion summary rebuilt from qc_cycle_completions.' -Data @{
            written = $true
            sheetPackageId = $packageId
            productionQcCompletedCount = $productionCount
            peerReviewCompletedCount = $peerCount
            independentCheckCompletedCount = $independentCount
        }
    } catch {
        return New-QCErrorResult -Code 'QC_CYCLE_SUMMARY_EXCEPTION' -Message $_.Exception.Message -Data @{ written = $false; sheetPackageId = $packageId }
    }
}

function Upsert-SheetPackageQcPdf {
    <#
    .SYNOPSIS
    Upserts one active lane QC PDF row for a sheet package (one active row per package + qc_process_type).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][guid]$SheetPackageId,
        [Parameter(Mandatory)][string]$QcProcessType,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$FolderPath = '',
        [string]$CurrentPwState = '',
        [string]$PreviousPwState = '',
        [string]$AssignedReviewerEmail = '',
        [string]$AssignedReviewerName = '',
        [Nullable[int]]$ReviewCycle = $null,
        [switch]$Required
    )
    if (-not (_QDB-IsEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'SHEET_PACKAGE_QC_PDF_SKIPPED' -Message 'Database telemetry is disabled.' -Data @{ written = $false }
    }
    $processType = ([string]$QcProcessType).Trim().ToLowerInvariant()
    if ($processType -notin @('production', 'check', 'review')) {
        return New-QCFailureResult -Code 'SHEET_PACKAGE_QC_PDF_INVALID_TYPE' -Message 'qc_process_type must be production, check, or review.' -Data @{ qcProcessType = $processType }
    }
    $parsedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    if (-not $parsedGuid) {
        if ($Required) {
            return New-QCFailureResult -Code 'QC_PACKAGE_QC_PDF_REQUIRED_MISSING' -Message 'Required lane QC PDF GUID is missing.' -Data @{
                sheetPackageId = $SheetPackageId; qcProcessType = $processType; documentName = $DocumentName
            }
        }
        return New-QCSuccessResult -Code 'QC_PACKAGE_QC_PDF_MISSING_OPTIONAL' -Message 'Optional lane QC PDF not resolved.' -Data @{
            written = $false; sheetPackageId = $SheetPackageId; qcProcessType = $processType
        }
    }
    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath
    try {
        $sql = @"
UPDATE sheet_package_qc_pdfs
SET is_active = 0, updated_at = SYSDATETIMEOFFSET()
WHERE sheet_package_id = @sheetPackageId
  AND qc_process_type = @qcProcessType
  AND is_active = 1
  AND document_guid <> @docGuid;

MERGE sheet_package_qc_pdfs AS tgt
USING (
    SELECT @sheetPackageId AS sheet_package_id, @qcProcessType AS qc_process_type, @docGuid AS document_guid
) AS src
ON tgt.sheet_package_id = src.sheet_package_id
   AND tgt.qc_process_type = src.qc_process_type
   AND tgt.document_guid = src.document_guid
   AND tgt.is_active = 1
WHEN MATCHED THEN UPDATE SET
    document_name = @docName,
    folder_path = COALESCE(@folderPath, tgt.folder_path),
    current_pw_state = COALESCE(@currentPwState, tgt.current_pw_state),
    previous_pw_state = COALESCE(@previousPwState, tgt.previous_pw_state),
    assigned_reviewer_email = COALESCE(@assignedReviewerEmail, tgt.assigned_reviewer_email),
    assigned_reviewer_name = COALESCE(@assignedReviewerName, tgt.assigned_reviewer_name),
    review_cycle = COALESCE(@reviewCycle, tgt.review_cycle),
    last_seen_at = SYSDATETIMEOFFSET(),
    updated_at = SYSDATETIMEOFFSET(),
    is_active = 1
WHEN NOT MATCHED THEN INSERT (
    sheet_package_id, qc_process_type, document_guid, document_name, folder_path,
    current_pw_state, previous_pw_state, assigned_reviewer_email, assigned_reviewer_name,
    review_cycle, last_seen_at, is_active
) VALUES (
    @sheetPackageId, @qcProcessType, @docGuid, @docName, @folderPath,
    @currentPwState, @previousPwState,
    @assignedReviewerEmail, @assignedReviewerName,
    @reviewCycle, SYSDATETIMEOFFSET(), 1
);
"@
        $params = @{
            sheetPackageId = $SheetPackageId
            qcProcessType = $processType
            docGuid = $parsedGuid
            docName = [string]$DocumentName
            folderPath = if ($folder) { $folder } else { [DBNull]::Value }
            currentPwState = if ($CurrentPwState) { $CurrentPwState } else { [DBNull]::Value }
            previousPwState = if ($PreviousPwState) { $PreviousPwState } else { [DBNull]::Value }
            assignedReviewerEmail = if ($AssignedReviewerEmail) { $AssignedReviewerEmail } else { [DBNull]::Value }
            assignedReviewerName = if ($AssignedReviewerName) { $AssignedReviewerName } else { [DBNull]::Value }
            reviewCycle = if ($null -ne $ReviewCycle) { $ReviewCycle } else { [DBNull]::Value }
        }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $dbRes.IsSuccess) {
            return New-QCErrorResult -Code 'SHEET_PACKAGE_QC_PDF_WRITE_FAILED' -Message $dbRes.Message -Data @{
                sheetPackageId = $SheetPackageId; qcProcessType = $processType
            }
        }
        return New-QCSuccessResult -Code 'QC_PACKAGE_QC_PDF_UPSERTED' -Message 'Lane QC PDF upserted.' -Data @{
            written = $true
            sheetPackageId = $SheetPackageId
            qcProcessType = $processType
            documentGuid = $parsedGuid
            documentName = $DocumentName
            currentPwState = $CurrentPwState
        }
    } catch {
        return New-QCErrorResult -Code 'SHEET_PACKAGE_QC_PDF_EXCEPTION' -Message $_.Exception.Message -Data @{
            sheetPackageId = $SheetPackageId; qcProcessType = $processType
        }
    }
}

function Update-SheetPackageQcPdfLaneState {
    <#
    .SYNOPSIS
    Updates workflow state for one active lane QC PDF row (sheet_package_id + qc_process_type + document_guid).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [string]$CurrentPwState = '',
        [string]$PreviousPwState = '',
        [Parameter(Mandatory)][string]$QcProcessType,
        [guid]$SheetPackageId = [guid]::Empty
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return $null }
    $parsedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    if (-not $parsedGuid) { return $null }
    $processType = ([string]$QcProcessType).Trim().ToLowerInvariant()
    if ($processType -notin @('production', 'check', 'review')) { return $null }
    if (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            if ($CurrentPwState) { $CurrentPwState = Format-QCWorkflowStateName -StateName $CurrentPwState -Config $Config }
            if ($PreviousPwState) { $PreviousPwState = Format-QCWorkflowStateName -StateName $PreviousPwState -Config $Config }
        } catch { }
    }
    try {
        $sql = @"
UPDATE sheet_package_qc_pdfs
SET previous_pw_state = CASE WHEN @previousPwState IS NOT NULL THEN @previousPwState ELSE current_pw_state END,
    current_pw_state = COALESCE(@currentPwState, current_pw_state),
    updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
  AND qc_process_type = @qcProcessType
  AND is_active = 1
"@
        if ($SheetPackageId -ne [guid]::Empty) {
            $sql += "  AND sheet_package_id = @sheetPackageId`n"
        }
        $params = @{
            docGuid = $parsedGuid
            qcProcessType = $processType
            currentPwState = if ($CurrentPwState) { $CurrentPwState } else { [DBNull]::Value }
            previousPwState = if ($PreviousPwState) { $PreviousPwState } else { [DBNull]::Value }
        }
        if ($SheetPackageId -ne [guid]::Empty) { $params['sheetPackageId'] = $SheetPackageId }
        $dbRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql $sql -Parameters $params
        if ($dbRes.IsSuccess -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
            $unchanged = [string]::IsNullOrWhiteSpace($CurrentPwState)
            Write-QCJsonLog -Level 'Information' -Code $(if ($unchanged) { 'QC_LANE_STATE_UNCHANGED' } else { 'QC_LANE_STATE_UPDATED' }) `
                -Message $(if ($unchanged) { 'Lane QC PDF state unchanged.' } else { 'Lane QC PDF state updated.' }) -Data @{
                documentGuid = $parsedGuid
                qcProcessType = $QcProcessType
                previousPwState = $PreviousPwState
                currentPwState = $CurrentPwState
            } | Out-Null
        }
        return $dbRes
    } catch { return $null }
}

function Sync-SheetPackageLaneQcPdfsFromMembers {
    <#
    .SYNOPSIS
    Discovers lane QC PDFs from resolved sheet members and upserts sheet_package_qc_pdfs rows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SheetStem,
        [Parameter(Mandatory)][array]$Members,
        [hashtable]$StateByGuid = @{},
        [string]$ActiveQcProcessType = '',
        [switch]$RequireActiveLane
    )
    if (-not (_QDB-IsEnabled -Config $Config)) { return @() }
    if ([string]::IsNullOrWhiteSpace($SheetStem)) { return @() }

    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath
    $ensure = Ensure-SheetPackage -Config $Config -FolderPath $folder -SheetStem $SheetStem -DocumentRole 'sheet_pdf'
    if (-not $ensure.IsSuccess -or -not $ensure.Data.sheetPackageId) { return @() }
    $packageId = [guid]$ensure.Data.sheetPackageId

    $laneTypes = @('production', 'check', 'review')
    $suffixMap = @{ production = '-prod.pdf'; check = '-chk.pdf'; review = '-rev.pdf' }
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($lane in $laneTypes) {
        $expectedSuffix = $suffixMap[$lane]
        $expectedName = ($SheetStem + $expectedSuffix).ToLowerInvariant()
        $member = @($Members | Where-Object {
            $dn = [string]$_.documentName
            $dn -and ($dn.ToLowerInvariant() -eq $expectedName)
        } | Select-Object -First 1)

        if ($member.Count -eq 0 -or -not $member[0]) {
            if ($RequireActiveLane -and $ActiveQcProcessType -eq $lane) {
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Level 'Warning' -Code 'QC_PACKAGE_QC_PDF_REQUIRED_MISSING' `
                        -Message 'Required lane QC PDF missing during package discovery.' -Data @{
                        sheetPackageId = $packageId; qcProcessType = $lane; folderPath = $folder; sheetStem = $SheetStem
                    } | Out-Null
                }
            } elseif (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'QC_PACKAGE_QC_PDF_MISSING_OPTIONAL' `
                    -Message 'Optional lane QC PDF not present in package discovery.' -Data @{
                    sheetPackageId = $packageId; qcProcessType = $lane; folderPath = $folder; sheetStem = $SheetStem
                } | Out-Null
            }
            continue
        }

        $m = $member[0]
        $dg = [string]$m.documentGuid
        $dn = [string]$m.documentName
        $pwState = ''
        if ($dg -and $StateByGuid -and $StateByGuid.ContainsKey($dg.ToLowerInvariant())) {
            $pwState = [string]$StateByGuid[$dg.ToLowerInvariant()]
        } elseif ($m.document -and (Get-Command -Name '_PWD-GetWorkflowStateFromDocumentRow' -ErrorAction SilentlyContinue)) {
            $pwState = [string](_PWD-GetWorkflowStateFromDocumentRow -DocRow $m.document)
        }

        $upsert = Upsert-SheetPackageQcPdf -Config $Config -SheetPackageId $packageId `
            -QcProcessType $lane -DocumentGuid $dg -DocumentName $dn -FolderPath $folder `
            -CurrentPwState $pwState -Required:($RequireActiveLane -and $ActiveQcProcessType -eq $lane)
        $results.Add($upsert) | Out-Null

        # Mirror lane GUID columns on sheet_packages for backward-compatible reporting.
        $colGuid = switch ($lane) {
            'production' { 'qc_pdf_guid' }
            'check' { 'qc_chk_pdf_guid' }
            'review' { 'qc_rev_pdf_guid' }
        }
        $colName = switch ($lane) {
            'production' { 'qc_pdf_name' }
            'check' { 'qc_chk_pdf_name' }
            'review' { 'qc_rev_pdf_name' }
        }
        $laneGuid = _QDB-TryParseDocumentGuid -DocumentGuid $dg
        if ($upsert.IsSuccess -and $laneGuid) {
            try {
                [void](Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_packages SET $colGuid = @docGuid, $colName = @docName, last_updated_at = SYSDATETIMEOFFSET()
WHERE sheet_package_id = @sheetPackageId
"@ -Parameters @{
                    docGuid = $laneGuid
                    docName = $dn
                    sheetPackageId = $packageId
                })
            } catch { }
        }
    }

    return @($results)
}

function Remove-QCLaneQcPdfRegistryRecords {
    <#
    .SYNOPSIS
    Purges registry/index rows for a deleted lane QC PDF (*-prod/-chk/-rev.pdf) and records a QC workflow event.
    .DESCRIPTION
    Called on DOCUMENT_DELETE audit events. Deactivates sheet_package_qc_pdfs, removes sheet_index and
    sheet_documents rows for the deleted GUID, and clears lane columns on sheet_packages when they still
    point at the deleted GUID. Does not touch transition_events, qc_workflow_events history from prior
    transitions, document_state_history, notification_log, or processing_jobs (except the new delete event).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$FolderPath = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$LastAuditEventAt = '',
        [switch]$DryRun
    )

    $docName = ([string]$DocumentName).Trim()
    $lane = _QDB-GetQcPdfLaneFromDocumentName -DocumentName $docName
    if (-not $lane) {
        return New-QCSuccessResult -Code 'LANE_PDF_DELETE_SKIPPED' -Message 'Not a lane QC PDF filename; registry cleanup skipped.' -Data @{
            skipped = $true; reason = 'not_lane_pdf'; documentName = $docName
        }
    }

    $parsedGuid = _QDB-TryParseDocumentGuid -DocumentGuid $DocumentGuid
    if (-not $parsedGuid) {
        return New-QCFailureResult -Code 'LANE_PDF_DELETE_GUID_INVALID' -Message 'document_guid is not a valid GUID.' -Data @{ documentGuid = $DocumentGuid }
    }
    $guidStr = $parsedGuid.ToString()
    $folder = _QDB-NormalizeTelemetryPath -Path $FolderPath

    $packageId = $null
    try { $packageId = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $guidStr } catch { }
    if (-not $packageId -and $folder) {
        try {
            $resolved = Resolve-SheetPackageFromDocument -DocumentGuid $guidStr -DocumentName $docName -FolderPath $folder
            if ($resolved.isSheetPackageMember) {
                $packageId = Resolve-SheetPackageIdForSheetGroup -Config $Config -FolderPath $folder `
                    -SheetStem $resolved.sheetStem -DocumentGuid $guidStr -DocumentName $docName
            }
        } catch { }
    }

    $prevState = ''
    if (_QDB-IsEnabled -Config $Config) {
        try {
            $stRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 COALESCE(q.current_pw_state, q.pw_state_name, q.previous_pw_state) AS last_state
FROM (
    SELECT current_pw_state, previous_pw_state, CAST(NULL AS NVARCHAR(100)) AS pw_state_name
    FROM sheet_package_qc_pdfs WHERE document_guid = @docGuid
    UNION ALL
    SELECT CAST(NULL AS NVARCHAR(100)), CAST(NULL AS NVARCHAR(100)), pw_state_name
    FROM sheet_index WHERE document_guid = @docGuid
) q
WHERE COALESCE(q.current_pw_state, q.pw_state_name, q.previous_pw_state) IS NOT NULL
"@ -Parameters @{ docGuid = $parsedGuid }
            if ($stRes.IsSuccess -and $stRes.Data.table -and $stRes.Data.table.Rows.Count -gt 0) {
                $cell = $stRes.Data.table.Rows[0].last_state
                if ($cell -isnot [DBNull] -and -not [string]::IsNullOrWhiteSpace([string]$cell)) {
                    $prevState = [string]$cell
                }
            }
        } catch { }
    }

    $counts = @{
        sheetPackageQcPdfsDeactivated = 0
        sheetIndexDeleted             = 0
        sheetIndexQcLinksCleared      = 0
        sheetPackagesLaneCleared      = 0
        sheetDocumentsDeleted         = 0
    }
    $writesAllowed = (Test-QCDatabaseWritesAllowed -Config $Config) -and -not $DryRun

    if ($writesAllowed) {
        $deactRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_package_qc_pdfs
SET is_active = 0, updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid AND is_active = 1
"@ -Parameters @{ docGuid = $parsedGuid }
        if (-not $deactRes.IsSuccess) {
            return New-QCErrorResult -Code 'LANE_PDF_DELETE_DEACTIVATE_FAILED' -Message $deactRes.Message -Data @{
                documentGuid = $guidStr; qcProcessType = $lane
            }
        }
        if ($deactRes.Data.rowsAffected) { $counts.sheetPackageQcPdfsDeactivated = [int]$deactRes.Data.rowsAffected }

        $idxDelRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
DELETE FROM sheet_index WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $parsedGuid }
        if (-not $idxDelRes.IsSuccess) {
            return New-QCErrorResult -Code 'LANE_PDF_DELETE_INDEX_FAILED' -Message $idxDelRes.Message -Data @{ documentGuid = $guidStr }
        }
        if ($idxDelRes.Data.rowsAffected) { $counts.sheetIndexDeleted = [int]$idxDelRes.Data.rowsAffected }

        $idxLinkRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index
SET qc_pdf_guid = NULL, qc_pdf_name = NULL, last_updated_at = SYSDATETIMEOFFSET()
WHERE qc_pdf_guid = @guidStr
"@ -Parameters @{ guidStr = $guidStr }
        if (-not $idxLinkRes.IsSuccess) {
            return New-QCErrorResult -Code 'LANE_PDF_DELETE_INDEX_LINK_FAILED' -Message $idxLinkRes.Message -Data @{ documentGuid = $guidStr }
        }
        if ($idxLinkRes.Data.rowsAffected) { $counts.sheetIndexQcLinksCleared = [int]$idxLinkRes.Data.rowsAffected }

        $pkgClrRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_packages SET
    qc_pdf_guid = CASE WHEN qc_pdf_guid = @docGuid THEN NULL ELSE qc_pdf_guid END,
    qc_pdf_name = CASE WHEN qc_pdf_guid = @docGuid THEN NULL ELSE qc_pdf_name END,
    qc_chk_pdf_guid = CASE WHEN qc_chk_pdf_guid = @docGuid THEN NULL ELSE qc_chk_pdf_guid END,
    qc_chk_pdf_name = CASE WHEN qc_chk_pdf_guid = @docGuid THEN NULL ELSE qc_chk_pdf_name END,
    qc_rev_pdf_guid = CASE WHEN qc_rev_pdf_guid = @docGuid THEN NULL ELSE qc_rev_pdf_guid END,
    qc_rev_pdf_name = CASE WHEN qc_rev_pdf_guid = @docGuid THEN NULL ELSE qc_rev_pdf_name END,
    last_updated_at = SYSDATETIMEOFFSET()
WHERE qc_pdf_guid = @docGuid OR qc_chk_pdf_guid = @docGuid OR qc_rev_pdf_guid = @docGuid
"@ -Parameters @{ docGuid = $parsedGuid }
        if (-not $pkgClrRes.IsSuccess) {
            return New-QCErrorResult -Code 'LANE_PDF_DELETE_PACKAGE_FAILED' -Message $pkgClrRes.Message -Data @{ documentGuid = $guidStr }
        }
        if ($pkgClrRes.Data.rowsAffected) { $counts.sheetPackagesLaneCleared = [int]$pkgClrRes.Data.rowsAffected }

        $sdDelRes = Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
DELETE FROM sheet_documents WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $parsedGuid }
        if (-not $sdDelRes.IsSuccess) {
            return New-QCErrorResult -Code 'LANE_PDF_DELETE_SHEET_DOC_FAILED' -Message $sdDelRes.Message -Data @{ documentGuid = $guidStr }
        }
        if ($sdDelRes.Data.rowsAffected) { $counts.sheetDocumentsDeleted = [int]$sdDelRes.Data.rowsAffected }
    }

    $jobId = ''
    if ($null -ne $AuditEventId -and $AuditEventId -gt 0) {
        $jobId = 'qc_lane_delete_{0}' -f $AuditEventId
    }

    $payloadObj = @{
        auditAction     = 'DOCUMENT_DELETE'
        documentName    = $docName
        folderPath      = $folder
        qcProcessType   = $lane
        auditEventId    = $AuditEventId
        lastAuditEventAt = $LastAuditEventAt
        registryCleanup = $counts
        dryRun          = [bool]$DryRun
        writesApplied   = [bool]$writesAllowed
    }
    $payloadJson = ''
    try { $payloadJson = ($payloadObj | ConvertTo-Json -Compress) } catch { }

    $wfParams = @{
        Config       = $Config
        DocumentId   = $guidStr
        JobId        = $jobId
        EventType    = 'DOCUMENT_DELETE'
        PreviousPwState = $prevState
        TargetPwState   = ''
        DecisionCode    = 'QC_LANE_PDF_REGISTRY_PURGED'
        QcReviewType    = $lane
        PayloadJson     = $payloadJson
    }
    if ($null -ne $packageId) { $wfParams['SheetPackageId'] = $packageId }
    if ($DryRun -or -not $writesAllowed) { $wfParams['PlannedOnly'] = $true }

    $wfRes = $null
    if (Get-Command -Name 'Write-QCWorkflowEventRow' -ErrorAction SilentlyContinue) {
        $wfRes = Write-QCWorkflowEventRow @wfParams
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_LANE_PDF_DELETED' `
            -Message 'Lane QC PDF DOCUMENT_DELETE processed; registry purged and QC workflow event recorded.' -Data @{
            documentGuid   = $guidStr
            documentName   = $docName
            folderPath     = $folder
            qcProcessType  = $lane
            auditEventId   = $AuditEventId
            sheetPackageId = if ($packageId) { $packageId.ToString() } else { '' }
            registryCleanup = $counts
            workflowEventWritten = if ($wfRes) { [bool]$wfRes.IsSuccess } else { $false }
            dryRun         = [bool]$DryRun
        } | Out-Null
    }

    return New-QCSuccessResult -Code 'LANE_PDF_DELETE_PROCESSED' -Message 'Lane QC PDF delete registry cleanup completed.' -Data @{
        written         = [bool]$writesAllowed
        documentGuid    = $guidStr
        documentName    = $docName
        qcProcessType   = $lane
        sheetPackageId  = if ($packageId) { $packageId.ToString() } else { '' }
        registryCleanup = $counts
        workflowEvent   = if ($wfRes) { $wfRes.Code } else { '' }
    }
}

Export-ModuleMember -Function Test-QCDatabaseEnabled, Test-QCDatabaseWritesAllowed, Test-QCSheetIndexFolderPath, Get-QCDatabaseConnection, Invoke-QCDatabaseQuery, Invoke-QCDatabaseNonQuery, Invoke-QCDatabaseScalar, Invoke-QCDatabaseBatch, New-QCDatabaseSession, Invoke-QCDatabaseNonQueryWithConnection, Invoke-QCDatabaseScalarWithConnection, Initialize-QCDatabaseSchema, Get-QCProcessingJobType, New-QCStateChangeJobId, Write-QCStateChangeJobTelemetry, Write-QCAuditEventRows, Write-QCJobTelemetry, Write-QCPollRunTelemetry, Write-QCDocumentStateHistoryRow, Write-QCWorkflowEventRow, Write-QCTransitionEvent, Ensure-QCTransitionEvent, Test-QCTransitionEventNotificationSent, Update-QCTransitionEventNotification, Get-QCTransitionEventActor, Get-QCAuditEventActor, Write-QCNotificationTelemetry, Resolve-SheetPackageFromDocument, Get-SheetPackageIdForDocument, Resolve-SheetPackageIdForSheetGroup, Resolve-QCCycleCompletionSheetPackageId, Ensure-SheetPackage, Write-SheetDocument, Upsert-SheetPackageQcPdf, Update-SheetPackageQcPdfLaneState, Sync-SheetPackageLaneQcPdfsFromMembers, Build-SheetPackageBackfillPlan, Write-QCSheetIndex, Write-QCSheetIndexBatch, Update-QCSheetIndexPwStateName, Get-QCSheetIndexCycle, Update-QCSheetIndexCycle, Resolve-QCSheetQcPdfGuid, Sync-QCLaneQcPdfGuidFromProjectWise, Update-QCSheetQcPdf, Remove-QCLaneQcPdfRegistryRecords, Get-QCReviewTypeBucket, Ensure-QCCycleCompletion, Update-QCSheetCycleCompletionSummary, Get-QCPWUnresolvedUserNumbers, Get-QCPWUserIdentity, Write-QCPWUserDirectory, Get-QCDocumentFolderCache, Get-QCNewerSheetDocumentStateAuditEvent, Get-QCUnprocessedAuditEvents, Update-QCAuditEventsResolvedFolders, Mark-QCAuditEventsProcessed, Upsert-QCDocumentActivityFolder, Get-QCWatcherStateValue, Set-QCWatcherStateValue, Get-QCAuditWatermarkUtc, Set-QCAuditWatermarkUtc, Get-QCPwDocumentCacheBatch, Set-QCPwDocumentCacheEntry, Get-QCPwFolderGuidCache, Get-QCPwFolderCacheBatch, Set-QCPwFolderCacheEntry, Update-QCProcessingJobCheckpoint, Update-QCProcessingJobHeartbeat
