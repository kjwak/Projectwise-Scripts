Import-Module (Join-Path $PSScriptRoot '..\Database\Core.Database.psm1') -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot '..\Core\Core.Results.psm1') -DisableNameChecking

function Get-QCOpsExplorerDefinitions {
    @(
        @{ label = 'Job history'; name = 'processing_jobs'; time = 'created_at'; columns = 'created_at,job_type,status,source_folder,source_path,error_code,duration_ms,job_id' }
        @{ label = 'Job timeline'; name = 'v_mcp_job_timeline'; time = 'ts'; columns = 'ts,process_name,level,code,message,job_id,document_guid' }
        @{ label = 'Recent errors'; name = 'v_mcp_recent_errors'; time = 'ts'; columns = 'ts,level,process_name,code,message,job_id' }
        @{ label = 'Sheet package status'; name = 'v_sheet_package_status'; columns = 'sheet_number,sheet_stem,stem_name,folder_path,pw_state_name,qc_process_type,sheet_package_id' }
        @{ label = 'Lane PDF matrix'; name = 'v_sheet_package_qc_pdf_matrix'; columns = 'sheet_number,folder_path,production_qc_pdf_name,production_state,review_qc_pdf_name,review_state,check_qc_pdf_name,check_state' }
        @{ label = 'Audit history'; name = 'v_audit_events_with_user'; time = 'captured_at'; columns = 'captured_at,pw_acttime,pw_action_name,resolved_folder,pw_itemname,pw_username,pw_objguid,processed' }
        @{ label = 'Poll and reconciliation runs'; name = 'poll_runs'; time = 'started_at'; columns = 'started_at,is_reconciliation,duration_ms,events_fetched,events_relevant,jobs_enqueued,error_message,watermark_after' }
        @{ label = 'Watcher state and schedule'; name = 'watcher_state'; columns = 'state_key,state_value,updated_at' }
        @{ label = 'Notification history'; name = 'notification_log'; time = 'sent_at'; columns = 'sent_at,success,event_type,document_name,folder_path,error_message' }
        @{ label = 'Automation events'; name = 'automation_events'; time = 'ts'; columns = 'ts,level,process_name,code,message,job_id,document_guid' }
    )
}

function Get-QCOpsExplorerCatalog {
    param([hashtable]$Config)
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'OPS_SQL_DISABLED' -Message 'Database is disabled in this profile.'
    }
    $sql = @'
SELECT s.name AS schema_name, o.name AS object_name, o.type AS object_type, c.name AS column_name,
       TYPE_NAME(c.system_type_id) AS data_type, c.max_length, ISNULL(pk.key_ordinal, 0) AS key_ordinal
FROM sys.objects o
JOIN sys.schemas s ON s.schema_id = o.schema_id
JOIN sys.columns c ON c.object_id = o.object_id
OUTER APPLY (
    SELECT ic.key_ordinal FROM sys.indexes i
    JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
    WHERE i.object_id = o.object_id AND i.is_primary_key = 1 AND ic.column_id = c.column_id
) pk
WHERE o.type IN ('U', 'V') AND o.is_ms_shipped = 0
  AND HAS_PERMS_BY_NAME(QUOTENAME(s.name) + '.' + QUOTENAME(o.name), 'OBJECT', 'SELECT') = 1
ORDER BY s.name, o.name, c.column_id
'@
    $result = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -CommandTimeout 15
    if (-not $result.IsSuccess) { return $result }
    $definitions = @{}
    foreach ($definition in Get-QCOpsExplorerDefinitions) { $definitions[$definition.name] = $definition }
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $result.Data.table.Rows | Group-Object schema_name, object_name) {
        $first = $group.Group[0]
        $name = [string]$first.object_name
        $schema = [string]$first.schema_name
        $definition = if ($schema -eq 'dbo') { $definitions[$name] } else { $null }
        $columns = @($group.Group | ForEach-Object {
            [pscustomobject]@{
                name = [string]$_.column_name
                type = [string]$_.data_type
                sortable = ($_.data_type -in @('bigint','int','smallint','tinyint','bit','float','real','decimal','numeric','money','smallmoney','date','datetime','datetime2','datetimeoffset','smalldatetime','time','char','varchar','nchar','nvarchar','uniqueidentifier','binary','varbinary','timestamp') -and $_.max_length -ne -1)
                key = [int]$_.key_ordinal
            }
        })
        $names = @($columns.name)
        $dates = @($columns | Where-Object { $_.type -in @('date','datetime','datetime2','datetimeoffset','smalldatetime') } | ForEach-Object { $_.name })
        $time = @(@($definition.time,'created_at','started_at','sent_at','captured_at','completed_at','occurred_at','event_at','ts_utc','ts','updated_at','last_seen_at','updated_utc','captured_utc','created_utc') |
            Where-Object { $_ -and $_ -in $dates } | Select-Object -First 1)
        if (-not $time.Count) { $time = @($dates | Select-Object -First 1) }
        $defaults = @()
        if ($definition) { $defaults = @($definition.columns.Split(',') | Where-Object { $_ -in $names }) }
        if ($defaults.Count -lt 3) { $defaults = @($names | Select-Object -First 10) }
        $items.Add([pscustomobject]@{
            id = "$schema.$name"; schema = $schema; name = $name; columns = $columns
            label = if ($definition) { "$($definition.label) [$schema.$name]" } else { "$schema.$name" }
            curated = [bool]$definition; defaults = $defaults; time = ($time -join '')
            folder = (@('folder_path','source_folder','resolved_folder','project_path','watch_folder','pw_folder_path') | Where-Object { $_ -in $names } | Select-Object -First 1)
        })
    }
    return New-QCSuccessResult -Code 'OPS_SQL_CATALOG' -Message "$($items.Count) queryable datasets." -Data @{
        items = @($items.ToArray() | Sort-Object @{ Expression = { -not $_.curated } }, label)
    }
}

function _OpsSqlIdentifier([string]$Name) { '[' + $Name.Replace(']', ']]') + ']' }
function _OpsSqlPattern([string]$Value) { $Value.Replace('[','[[]').Replace('%','[%]').Replace('_','[_]') }

function New-QCOpsExplorerQuery {
    param(
        [Parameter(Mandatory)]$Dataset, [hashtable]$Filter = @{},
        [string]$Sort, [switch]$Ascending, [int]$Page = 0, [int]$PageSize = 100, [switch]$Export,
        [string[]]$Projection
    )
    $names = @($Dataset.columns.name)
    $sortable = @($Dataset.columns | Where-Object sortable | ForEach-Object { $_.name })
    if ($Sort -and $Sort -notin $sortable) { throw 'Choose a sortable column from this dataset.' }
    if (-not $Sort) { $Sort = $Dataset.time }
    $keys = @($Dataset.columns | Where-Object { $_.key -gt 0 -and $_.sortable } | Sort-Object key | ForEach-Object { $_.name })
    $order = @($Sort) + $keys
    if (-not $keys.Count) {
        $identity = @('id','event_id','job_id','sheet_package_id','document_guid','state_key') | Where-Object { $_ -in $sortable }
        $order += @($identity)
        if (-not $identity) { $order += $sortable }
    }
    $order = @($order | Where-Object { $_ -and $_ -in $sortable } | Select-Object -Unique)
    $direction = if ($Ascending) { 'ASC' } else { 'DESC' }
    $parts = @($order | ForEach-Object { (_OpsSqlIdentifier $_) + ' ' + $direction })
    $orderSql = if ($parts.Count) { $parts -join ', ' } else { '(SELECT NULL)' }
    $where = [System.Collections.Generic.List[string]]::new()
    $limit = if ($Export) { 5000 } else { [math]::Min(250, [math]::Max(1, $PageSize)) }
    $params = @{ skip = if ($Export) { 0L } else { [long][math]::Max(0, $Page) * $limit }; take = $limit + 1 }
    $folder = [string]$Filter.Folder
    if (-not $folder) { $folder = [string]$Filter.ProjectRoot }
    $folders = @(if ($folder) { $folder } else { $Filter.ProjectFolders | Where-Object { $_ } | Select-Object -Unique -First 100 })
    if ($folders.Count -and $Dataset.folder) {
        $folderColumn = _OpsSqlIdentifier $Dataset.folder
        $predicates = @(for ($i = 0; $i -lt $folders.Count; $i++) {
            $value = $folders[$i].TrimEnd('\','/')
            $suffix = if ($i -gt 0) { [string]$i } else { '' }
            $params["folder$suffix"] = $value
            $params["underFolder$suffix"] = (_OpsSqlPattern $value) + '\%'
            "($folderColumn = @folder$suffix OR $folderColumn LIKE @underFolder$suffix)"
        })
        $where.Add('(' + ($predicates -join ' OR ') + ')')
    }
    foreach ($bound in @('FromUtc','ToUtc')) {
        if ($Filter[$bound] -and $Dataset.time) {
            $comparison = if ($bound -eq 'FromUtc') { '>=' } else { '<=' }
            $where.Add("$(_OpsSqlIdentifier $Dataset.time) $comparison @$bound")
            $dateType = ($Dataset.columns | Where-Object name -eq $Dataset.time | Select-Object -First 1).type
            $params[$bound] = if ($dateType -eq 'datetimeoffset') { [datetimeoffset][datetime]$Filter[$bound] } else { [datetime]$Filter[$bound] }
        }
    }
    if ($Filter.Column -and ($Filter.Text -or $Filter.Exact)) {
        if ($Filter.Column -notin $names) { throw 'Choose a filter column from this dataset.' }
        $column = _OpsSqlIdentifier $Filter.Column
        if ($Filter.Exact) {
            $type = ($Dataset.columns | Where-Object name -eq $Filter.Column | Select-Object -First 1).type
            $filterValue = [string]$Filter.Text
            switch ($type) {
                'uniqueidentifier' { $filterValue = [guid]$filterValue; break }
                'bigint' { $filterValue = [long]$filterValue; break }
                { $_ -in @('int','smallint','tinyint') } { $filterValue = [int]$filterValue; break }
                { $_ -notin @('char','varchar','nchar','nvarchar') } { $column = "CONVERT(nvarchar(4000), $column)" }
            }
            $where.Add("$column = @filterText")
            $params.filterText = $filterValue
        } else {
            $where.Add("CONVERT(nvarchar(4000), $column) LIKE @filterText")
            $params.filterText = '%' + (_OpsSqlPattern ([string]$Filter.Text)) + '%'
        }
    }
    if ($Filter.ReconciliationOnly -and 'is_reconciliation' -in $names) { $where.Add('[is_reconciliation] = 1') }
    if ($Filter.AuditOnly -and 'is_reconciliation' -in $names) { $where.Add('[is_reconciliation] = 0') }
    if ($Filter.FailuresOnly -and 'success' -in $names) { $where.Add('[success] = 0') }
    $whereSql = if ($where.Count) { ' WHERE ' + ($where -join ' AND ') } else { '' }
    $object = (_OpsSqlIdentifier $Dataset.schema) + '.' + (_OpsSqlIdentifier $Dataset.name)
    $columns = '*'
    if ($Projection.Count) {
        if (@($Projection | Where-Object { $_ -notin $names }).Count) { throw 'Choose columns from this dataset.' }
        $columns = ($Projection | Select-Object -Unique | ForEach-Object { _OpsSqlIdentifier $_ }) -join ', '
    }
    @{
        sql = "SELECT $columns FROM $object$whereSql ORDER BY $orderSql OFFSET @skip ROWS FETCH NEXT @take ROWS ONLY"
        parameters = $params; limit = $limit
    }
}

function Invoke-QCOpsExplorerPage {
    param([hashtable]$Config, $Dataset, [hashtable]$Filter, [string]$Sort, [switch]$Ascending, [int]$Page, [int]$PageSize, [switch]$Export, [string[]]$Projection)
    $query = New-QCOpsExplorerQuery -Dataset $Dataset -Filter $Filter -Sort $Sort -Ascending:$Ascending -Page $Page -PageSize $PageSize -Export:$Export -Projection $Projection
    $result = Invoke-QCDatabaseQuery -Config $Config -Sql $query.sql -Parameters $query.parameters -CommandTimeout 20
    if ($result.IsSuccess) {
        $table = $result.Data.table
        $result.Data.hasMore = $table.Rows.Count -gt $query.limit
        if ($result.Data.hasMore) { $table.Rows.RemoveAt($table.Rows.Count - 1) }
        $result.Data.rowCount = $table.Rows.Count
    }
    return $result
}

function New-QCOpsCustomSelect {
    param([string]$Query)
    $query = $Query.Trim() -replace ';\s*$', ''
    if ($query.Length -gt 8000) { throw 'Queries are limited to 8,000 characters.' }
    $tokens = [regex]::Replace($query, "'(?:''|[^'])*'|\[(?:\]\]|[^\]])*\]|""(?:""""|[^""])*""", ' ')
    if ($tokens -notmatch '^SELECT\b' -or $tokens -match "['""\[\];]|--|/\*|\*/|\b(INTO|OPENQUERY|OPENROWSET|OPENDATASOURCE|EXEC(?:UTE)?|INSERT|UPDATE|DELETE|MERGE|CREATE|DROP|ALTER|TRUNCATE|GRANT|REVOKE|DENY|WAITFOR|OPTION|USE|SET|DBCC|BULK|NEXT\s+VALUE)\b") {
        throw 'Enter one SELECT without comments, INTO, external rowsets or query hints.'
    }
    # Constrain it to an expression. Statements cannot execute inside the derived table.
    return "SET TEXTSIZE 65536; SELECT TOP (251) * FROM ($query) AS qc_preview;"
}

function Invoke-QCOpsCustomSelect {
    param([string]$Query)
    $sql = New-QCOpsCustomSelect -Query $Query
    $connectionString = $env:QC_OPS_SQL_READONLY_CONNECTION_STRING
    if (-not $connectionString) {
        return New-QCFailureResult -Code 'OPS_SQL_READONLY_NOT_CONFIGURED' -Message 'Configure QC_OPS_SQL_READONLY_CONNECTION_STRING with a dedicated SQL principal granted SELECT only.'
    }
    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder($connectionString)
    $builder.ConnectTimeout = 5
    $result = Invoke-QCDatabaseQuery -Config @{ database = @{ connectionString = $builder.ConnectionString } } -Sql $sql -CommandTimeout 10
    if ($result.IsSuccess) {
        $result.Data.hasMore = $result.Data.table.Rows.Count -gt 250
        if ($result.Data.hasMore) { $result.Data.table.Rows.RemoveAt(250) }
        $result.Data.rowCount = $result.Data.table.Rows.Count
    }
    return $result
}

function Export-QCOpsExplorerCsv {
    param([System.Data.DataTable]$Table, [string[]]$Columns, [string]$Path)
    if (-not $Table.Rows.Count) { throw 'There are no matching rows to export.' }
    $Table.Rows | ForEach-Object {
        $values = [ordered]@{}
        foreach ($column in $Columns) {
            $value = $_[$column]
            if ($value -is [datetimeoffset]) { $value = $value.ToString('o') }
            if ($value -is [datetime]) { $value = $value.ToString('o') }
            if ($value -is [string] -and $value -match '^\s*[=+\-@]') { $value = "'" + $value }
            $values[$column] = $value
        }
        [pscustomobject]$values
    } | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

Export-ModuleMember -Function Get-QCOpsExplorerCatalog, New-QCOpsExplorerQuery, Invoke-QCOpsExplorerPage, New-QCOpsCustomSelect, Invoke-QCOpsCustomSelect, Export-QCOpsExplorerCsv
