# QC.DebugMcp.psm1
# Read-only diagnostics for QC workflow debugging (MCP / interactive use).

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Paths.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Database/Core.Database.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Telemetry.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'ProjectWise/PW.Connection.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'ProjectWise/PW.Discovery.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue

$script:QDM_Config = $null
$script:QDM_ColumnCache = @{}
$script:QDM_TableExistsCache = @{}

$script:QDM_SheetSearchTables = @(
    @{ table = 'sheet_index'; columns = @('document_name', 'document_guid', 'folder_path', 'sheet_package_id', 'pw_state_name') }
    @{ table = 'sheet_packages'; columns = @('sheet_stem', 'folder_path', 'sheet_package_id', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid', 'pw_state_name') }
    @{ table = 'sheet_documents'; columns = @('document_name', 'document_guid', 'sheet_package_id', 'document_role', 'pw_state_name') }
    @{ table = 'audit_events'; columns = @('pw_itemname', 'pw_objguid', 'resolved_folder', 'pw_textparam') }
    @{ table = 'transition_events'; columns = @('document_name', 'document_guid', 'folder_path', 'from_value', 'to_value') }
    @{ table = 'document_state_history'; columns = @('document_name', 'document_guid', 'folder_path', 'old_value', 'new_value') }
    @{ table = 'notification_log'; columns = @('document_name', 'document_guid', 'subject', 'folder_path') }
    @{ table = 'processing_jobs'; columns = @('source_path', 'source_folder', 'job_id', 'dedupe_key') }
    @{ table = 'qc_workflow_events'; columns = @('document_id', 'payload_json') }
)

$script:QDM_TimelineSources = @(
    @{
        source = 'audit_events'; table = 'audit_events'; time = 'captured_at'; id = 'id'
        document_name = 'pw_itemname'; document_guid = 'pw_objguid'; sheet_package_id = $null
        event_type = 'pw_action_name'; state_from = $null; state_to = $null; status = 'processed'; actor = 'pw_userno'
        detail_cols = @('pw_action', 'pw_acttime', 'resolved_folder', 'candidate_type', 'enqueued_job_id')
        where_cols = @('pw_itemname', 'pw_objguid', 'resolved_folder')
    }
    @{
        source = 'document_state_history'; table = 'document_state_history'; time = 'captured_at'; id = 'id'
        document_name = 'document_name'; document_guid = 'document_guid'; sheet_package_id = 'sheet_package_id'
        event_type = 'event_type'; state_from = 'old_value'; state_to = 'new_value'; status = $null; actor = 'changed_by_username'
        detail_cols = @('field_name', 'source_audit_id', 'transition_group_id', 'folder_path')
        where_cols = @('document_name', 'document_guid', 'folder_path')
    }
    @{
        source = 'transition_events'; table = 'transition_events'; time = 'detected_at'; id = 'id'
        document_name = 'document_name'; document_guid = 'document_guid'; sheet_package_id = 'sheet_package_id'
        event_type = 'transition_type'; state_from = 'from_value'; state_to = 'to_value'; status = 'notification_sent'; actor = 'changed_by_username'
        detail_cols = @('job_id', 'job_type', 'trigger_audit_id', 'notification_id', 'folder_path')
        where_cols = @('document_name', 'document_guid', 'folder_path')
    }
    @{
        source = 'qc_workflow_events'; table = 'qc_workflow_events'; time = 'created_utc'; id = 'event_id'
        document_name = $null; document_guid = 'document_id'; sheet_package_id = 'sheet_package_id'
        event_type = 'event_type'; state_from = 'previous_pw_state'; state_to = 'target_pw_state'; status = 'decision_code'; actor = $null
        detail_cols = @('transition_event_id', 'payload_json', 'processor_version')
        where_cols = @('document_id', 'payload_json')
    }
    @{
        source = 'notification_log'; table = 'notification_log'; time = 'sent_at'; id = 'id'
        document_name = 'document_name'; document_guid = 'document_guid'; sheet_package_id = 'sheet_package_id'
        event_type = 'event_type'; state_from = $null; state_to = $null; status = 'success'; actor = $null
        detail_cols = @('recipients', 'subject', 'dedupe_key', 'provider', 'error_message', 'transition_id')
        where_cols = @('document_name', 'document_guid', 'subject', 'folder_path')
    }
    @{
        source = 'processing_jobs'; table = 'processing_jobs'; time = 'created_at'; id = 'id'
        document_name = 'source_path'; document_guid = $null; sheet_package_id = 'sheet_package_id'
        event_type = 'job_type'; state_from = $null; state_to = $null; status = 'status'; actor = $null
        detail_cols = @('job_id', 'error_code', 'error_message', 'dedupe_key', 'trigger_audit_id', 'source_folder')
        where_cols = @('source_path', 'source_folder', 'job_id', 'dedupe_key')
    }
)

function Initialize-QCDebugMcpContext {
    [CmdletBinding()]
    param(
        [string]$AppSettingsPath = ''
    )

    if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
        $AppSettingsPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json'
    }
    if (-not (Test-Path -LiteralPath $AppSettingsPath)) {
        throw "appsettings not found: $AppSettingsPath"
    }
    $script:QDM_Config = Get-QCAppSettingsConfig -Path $AppSettingsPath
    if (-not (Test-QCDatabaseEnabled -Config $script:QDM_Config)) {
        throw 'database.enabled is false in appsettings; diagnostics require QC_Pipeline telemetry.'
    }

    $sqlServer = $env:PWQC_SQL_SERVER
    $sqlDatabase = $env:PWQC_SQL_DATABASE
    $sqlTrust = $env:PWQC_SQL_TRUST_CERT
    if ($sqlServer -or $sqlDatabase) {
        $conn = [string]$script:QDM_Config.database.connectionString
        if ($sqlServer) {
            if ($conn -match '(?i)Server\s*=') {
                $conn = [regex]::Replace($conn, '(?i)Server\s*=[^;]*', "Server=$sqlServer")
            } else {
                $conn = "Server=$sqlServer;$conn"
            }
        }
        if ($sqlDatabase) {
            if ($conn -match '(?i)Database\s*=') {
                $conn = [regex]::Replace($conn, '(?i)Database\s*=[^;]*', "Database=$sqlDatabase")
            } else {
                $conn = "$conn;Database=$sqlDatabase"
            }
        }
        if ($sqlTrust) {
            $trustVal = if ($sqlTrust -match '^(?i)(yes|true|1)$') { 'True' } else { 'False' }
            if ($conn -match '(?i)TrustServerCertificate\s*=') {
                $conn = [regex]::Replace($conn, '(?i)TrustServerCertificate\s*=[^;]*', "TrustServerCertificate=$trustVal")
            } else {
                $conn = "$conn;TrustServerCertificate=$trustVal"
            }
        }
        if (-not $script:QDM_Config.ContainsKey('database')) { $script:QDM_Config['database'] = @{} }
        $script:QDM_Config.database['connectionString'] = $conn
    }

    return $script:QDM_Config
}

function _QDM-Config {
    if (-not $script:QDM_Config) {
        throw 'Call Initialize-QCDebugMcpContext before using QC debug tools.'
    }
    return $script:QDM_Config
}

function _QDM-SerializeValue {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return $null }
    if ($Value -is [string] -or $Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return $Value }
    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return $Value.ToString('o') }
    if ($Value -is [guid]) { return $Value.ToString() }
    if ($Value -is [decimal]) { return [string]$Value }
    if ($Value -is [byte[]]) { return [System.Text.Encoding]::UTF8.GetString($Value) }
    return [string]$Value
}

function _QDM-RowsFromQuery {
    param(
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Parameters = @{}
    )
    $cfg = _QDM-Config
    $res = Invoke-QCDatabaseQuery -Config $cfg -Sql $Sql -Parameters $Parameters
    if (-not $res.IsSuccess) {
        throw $res.Message
    }
    $table = $res.Data.table
    if (-not $table -or $table.Rows.Count -eq 0) { return @() }
    $rows = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($dataRow in @($table.Rows)) {
        $h = @{}
        foreach ($col in $table.Columns) {
            $name = [string]$col.ColumnName
            $h[$name] = _QDM-SerializeValue -Value $dataRow[$name]
        }
        $rows.Add($h)
    }
    return @($rows)
}

function _QDM-BuildWarning {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Table = '',
        [string]$Column = ''
    )
    $w = @{ message = $Message }
    if ($Table) { $w['table'] = $Table }
    if ($Column) { $w['column'] = $Column }
    return $w
}

function _QDM-ToolResult {
    param(
        $Data,
        [object[]]$Warnings = @(),
        [string[]]$SourceTables = @(),
        [string[]]$QueryAssumptions = @()
    )
    return @{
        data = $Data
        warnings = @($Warnings)
        source_tables = @($SourceTables)
        query_assumptions = @($QueryAssumptions)
    }
}

function _QDM-SafeTopLimit {
    param([int]$Limit, [int]$MaxLimit = 500)
    if ($Limit -lt 1) { return $MaxLimit }
    return [Math]::Max(1, [Math]::Min($Limit, $MaxLimit))
}

function _QDM-TestTableExists {
    param([Parameter(Mandatory)][string]$TableName)
    $name = $TableName.Trim()
    if ($script:QDM_TableExistsCache.ContainsKey($name)) { return $script:QDM_TableExistsCache[$name] }
    $rows = _QDM-RowsFromQuery -Sql @"
SELECT 1 AS ok
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = @tableName
"@ -Parameters @{ tableName = $name }
    $exists = ($rows.Count -gt 0)
    $script:QDM_TableExistsCache[$name] = $exists
    return $exists
}

function _QDM-GetColumns {
    param([Parameter(Mandatory)][string]$TableName)
    $name = $TableName.Trim()
    if ($script:QDM_ColumnCache.ContainsKey($name)) { return @($script:QDM_ColumnCache[$name]) }
    $rows = _QDM-RowsFromQuery -Sql @"
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = @tableName
ORDER BY ORDINAL_POSITION
"@ -Parameters @{ tableName = $name }
    $cols = @($rows | ForEach-Object { [string]$_.COLUMN_NAME })
    $script:QDM_ColumnCache[$name] = $cols
    return $cols
}

function _QDM-SelectExistingColumns {
    param(
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string[]]$Requested
    )
    $existing = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($c in (_QDM-GetColumns -TableName $TableName)) { [void]$existing.Add($c) }
    return @($Requested | Where-Object { $existing.Contains($_) })
}

function _QDM-LikePattern {
    param([Parameter(Mandatory)][string]$Text)
    return "%$($Text.Trim())%"
}

function _QDM-BuildLikeWhere {
    param(
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$LikeParam
    )
    $usable = @(_QDM-SelectExistingColumns -TableName $TableName -Requested $Columns)
    if ($usable.Count -eq 0) { return $null }
    $clauses = @($usable | ForEach-Object { "CAST([$_] AS NVARCHAR(MAX)) LIKE @likeParam" })
    return @{ clause = ($clauses -join ' OR '); params = @{ likeParam = $LikeParam } }
}

function _QDM-TryParseGuid {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try {
        $g = [guid]::Parse($Value.Trim())
        if ($g -eq [guid]::Empty) { return $null }
        return $g.ToString()
    } catch { return $null }
}

function _QDM-ParseDocumentPath {
    <#
    .SYNOPSIS
    Parses a ProjectWise document path (pw:\\datasource\Documents\...\file.pdf) into telemetry keys.
    Paths ending in a folder segment (no file extension) are treated as folder-only lookups.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DocumentPath)

    $raw = $DocumentPath.Trim()
    $isFolderOnly = $raw -match '[\\/]\s*$'
    $normRes = Normalize-QCPath -Path $raw
    if (-not $normRes.IsSuccess) {
        return @{ error = $normRes.Message; raw_path = $raw }
    }

    $full = [string]$normRes.Data.path
    if ($isFolderOnly) {
        $folderPath = $full
        $folderRes = Normalize-QCDocumentsFolderPath -Path $folderPath
        if ($folderRes.IsSuccess) { $folderPath = [string]$folderRes.Data.path }
        return @{
            raw_path = $raw
            folder_path = $folderPath
            document_name = $null
            sheet_stem = $null
            is_folder_only = $true
        }
    }

    $documentName = [System.IO.Path]::GetFileName($full)
    $parent = [System.IO.Path]::GetDirectoryName($full)
    if ([string]::IsNullOrWhiteSpace($documentName)) {
        return @{ error = 'Could not extract document file name from path.'; raw_path = $raw }
    }

    $folderPath = $parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        $folderRes = Normalize-QCDocumentsFolderPath -Path $parent
        if ($folderRes.IsSuccess) { $folderPath = [string]$folderRes.Data.path }
    }

    $ext = [System.IO.Path]::GetExtension($documentName)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        $folderOnlyPath = $full
        $folderRes = Normalize-QCDocumentsFolderPath -Path $folderOnlyPath
        if ($folderRes.IsSuccess) { $folderOnlyPath = [string]$folderRes.Data.path }
        return @{
            raw_path = $raw
            folder_path = $folderOnlyPath
            document_name = $null
            sheet_stem = $null
            is_folder_only = $true
        }
    }

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($documentName)
    if ($stem -match '(?i)-(prod|chk|rev)$') { $stem = $stem -replace '(?i)-(prod|chk|rev)$', '' }

    return @{
        raw_path = $raw
        folder_path = $folderPath
        document_name = $documentName
        sheet_stem = $stem.ToLowerInvariant()
        is_folder_only = $false
    }
}

function _QDM-IngestFolderPathIdentity {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [System.Collections.Generic.HashSet[string]]$PackageIds,
        [System.Collections.Generic.HashSet[string]]$DocumentGuids,
        [System.Collections.Generic.HashSet[string]]$DocumentNames,
        [System.Collections.Generic.HashSet[string]]$SheetStems
    )

    if (_QDM-TestTableExists -TableName 'sheet_index') {
        $cols = _QDM-SelectExistingColumns -TableName 'sheet_index' -Requested @(
            'document_guid', 'document_name', 'sheet_package_id', 'sheet_stem', 'folder_path', 'qc_pdf_guid'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (200) $select FROM [sheet_index] WHERE folder_path = @folderPath ORDER BY last_updated_at DESC" -Parameters @{
                folderPath = $FolderPath
            }
            _QDM-IngestIdentityRows -Rows $rows -PackageIds $PackageIds -DocumentGuids $DocumentGuids -DocumentNames $DocumentNames -SheetStems $SheetStems
        }
    }

    if (_QDM-TestTableExists -TableName 'sheet_packages') {
        $cols = _QDM-SelectExistingColumns -TableName 'sheet_packages' -Requested @(
            'sheet_package_id', 'sheet_stem', 'folder_path', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (200) $select FROM [sheet_packages] WHERE folder_path = @folderPath ORDER BY last_updated_at DESC" -Parameters @{
                folderPath = $FolderPath
            }
            _QDM-IngestIdentityRows -Rows $rows -PackageIds $PackageIds -DocumentGuids $DocumentGuids -DocumentNames $DocumentNames -SheetStems $SheetStems
        }
    }

    if ($PackageIds.Count -gt 0 -and (_QDM-TestTableExists -TableName 'sheet_documents')) {
        $cols = _QDM-SelectExistingColumns -TableName 'sheet_documents' -Requested @(
            'document_guid', 'document_name', 'sheet_package_id', 'document_role'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $inList = (@($PackageIds) | ForEach-Object { "'$_'" }) -join ','
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (200) $select FROM [sheet_documents] WHERE CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)" -Parameters @{}
            _QDM-IngestIdentityRows -Rows $rows -PackageIds $PackageIds -DocumentGuids $DocumentGuids -DocumentNames $DocumentNames -SheetStems $SheetStems
        }
    }
}

function _QDM-IngestIdentityRows {
    param(
        [object[]]$Rows,
        [System.Collections.Generic.HashSet[string]]$PackageIds,
        [System.Collections.Generic.HashSet[string]]$DocumentGuids,
        [System.Collections.Generic.HashSet[string]]$DocumentNames,
        [System.Collections.Generic.HashSet[string]]$SheetStems
    )
    foreach ($row in @($Rows)) {
        foreach ($g in @('document_guid', 'pw_objguid', 'document_id', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid')) {
            if ($row[$g]) { [void]$DocumentGuids.Add([string]$row[$g]) }
        }
        if ($row.document_name) { [void]$DocumentNames.Add([string]$row.document_name) }
        if ($row.sheet_package_id) { [void]$PackageIds.Add([string]$row.sheet_package_id) }
        if ($row.sheet_stem) { [void]$SheetStems.Add([string]$row.sheet_stem) }
    }
}

function _QDM-GetLookupBoundParameters {
    param([hashtable]$Bound = @{})
    $out = @{}
    foreach ($key in @('SheetNumber', 'DocumentGuid', 'PackageId', 'SheetName', 'DocumentPath')) {
        if ($Bound.ContainsKey($key) -and $null -ne $Bound[$key] -and -not [string]::IsNullOrWhiteSpace([string]$Bound[$key])) {
            $out[$key] = $Bound[$key]
        }
    }
    return $out
}

function Resolve-QCDebugLookup {
    <#
    .SYNOPSIS
    Resolves sheet package IDs, document GUIDs, and names from sheet number, GUID, package ID, or document path.
    #>
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupText = ''
    $lookupType = ''
    $folderPath = $null
    $documentName = $null
    $sheetStem = $null
    $rawDocumentPath = $null
    $warnings = @()
    $packageIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $documentGuids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $documentNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $sheetStems = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    if (-not [string]::IsNullOrWhiteSpace($SheetName) -and [string]::IsNullOrWhiteSpace($SheetNumber)) {
        $SheetNumber = $SheetName
    }

    $parsedPkg = _QDM-TryParseGuid -Value $PackageId
    $parsedDoc = _QDM-TryParseGuid -Value $DocumentGuid

    if ($parsedPkg) {
        $lookupType = 'package_id'
        $lookupText = $parsedPkg
        [void]$packageIds.Add($parsedPkg)
        if (_QDM-TestTableExists -TableName 'sheet_packages') {
            $rows = _QDM-RowsFromQuery -Sql @"
SELECT sheet_package_id, sheet_stem, folder_path, dgn_guid, sheet_pdf_guid, qc_pdf_guid
FROM sheet_packages
WHERE CAST(sheet_package_id AS NVARCHAR(36)) = @pkg
"@ -Parameters @{ pkg = $parsedPkg }
            foreach ($row in $rows) {
                if ($row.sheet_stem) { [void]$sheetStems.Add([string]$row.sheet_stem) }
                foreach ($g in @('dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid')) {
                    if ($row[$g]) { [void]$documentGuids.Add([string]$row[$g]) }
                }
            }
        }
        if (_QDM-TestTableExists -TableName 'sheet_documents') {
            $rows = _QDM-RowsFromQuery -Sql @"
SELECT document_guid, document_name, document_role
FROM sheet_documents
WHERE CAST(sheet_package_id AS NVARCHAR(36)) = @pkg
"@ -Parameters @{ pkg = $parsedPkg }
            foreach ($row in $rows) {
                if ($row.document_guid) { [void]$documentGuids.Add([string]$row.document_guid) }
                if ($row.document_name) { [void]$documentNames.Add([string]$row.document_name) }
            }
        }
    }
    elseif ($parsedDoc) {
        $lookupType = 'document_guid'
        $lookupText = $parsedDoc
        [void]$documentGuids.Add($parsedDoc)
        $pkg = Get-SheetPackageIdForDocument -Config (_QDM-Config) -DocumentGuid $parsedDoc
        if ($pkg) { [void]$packageIds.Add($pkg.ToString()) }
        foreach ($table in @('sheet_index', 'sheet_documents', 'sheet_packages')) {
            if (-not (_QDM-TestTableExists -TableName $table)) { continue }
            $cols = _QDM-SelectExistingColumns -TableName $table -Requested @(
                'document_guid', 'document_name', 'sheet_package_id', 'sheet_stem', 'qc_pdf_guid', 'dgn_guid', 'sheet_pdf_guid'
            )
            if ($cols.Count -eq 0) { continue }
            $whereParts = @()
            $params = @{ docGuid = $parsedDoc }
            if ($cols -contains 'document_guid') { $whereParts += 'CAST(document_guid AS NVARCHAR(36)) = @docGuid' }
            if ($cols -contains 'qc_pdf_guid') { $whereParts += 'CAST(qc_pdf_guid AS NVARCHAR(36)) = @docGuid' }
            if ($cols -contains 'dgn_guid') { $whereParts += 'CAST(dgn_guid AS NVARCHAR(36)) = @docGuid' }
            if ($cols -contains 'sheet_pdf_guid') { $whereParts += 'CAST(sheet_pdf_guid AS NVARCHAR(36)) = @docGuid' }
            if ($whereParts.Count -eq 0) { continue }
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (50) $select FROM [$table] WHERE $($whereParts -join ' OR ')" -Parameters $params
            foreach ($row in $rows) {
                if ($row.document_name) { [void]$documentNames.Add([string]$row.document_name) }
                if ($row.sheet_package_id) { [void]$packageIds.Add([string]$row.sheet_package_id) }
                if ($row.sheet_stem) { [void]$sheetStems.Add([string]$row.sheet_stem) }
                foreach ($g in @('document_guid', 'qc_pdf_guid', 'dgn_guid', 'sheet_pdf_guid')) {
                    if ($row[$g]) { [void]$documentGuids.Add([string]$row[$g]) }
                }
            }
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($DocumentPath)) {
        $parsed = _QDM-ParseDocumentPath -DocumentPath $DocumentPath
        if ($parsed.error) {
            throw "Invalid document_path: $($parsed.error)"
        }
        if ($parsed.is_folder_only) {
            $lookupType = 'folder_path'
            $lookupText = [string]$parsed.raw_path
            $rawDocumentPath = [string]$parsed.raw_path
            $folderPath = [string]$parsed.folder_path
            _QDM-IngestFolderPathIdentity -FolderPath $folderPath -PackageIds $packageIds -DocumentGuids $documentGuids `
                -DocumentNames $documentNames -SheetStems $sheetStems
        } else {
        $lookupType = 'document_path'
        $lookupText = [string]$parsed.raw_path
        $rawDocumentPath = [string]$parsed.raw_path
        $folderPath = [string]$parsed.folder_path
        $documentName = [string]$parsed.document_name
        $sheetStem = [string]$parsed.sheet_stem
        [void]$documentNames.Add($documentName)
        [void]$sheetStems.Add($sheetStem)

        if (_QDM-TestTableExists -TableName 'sheet_index') {
            $cols = _QDM-SelectExistingColumns -TableName 'sheet_index' -Requested @(
                'document_guid', 'document_name', 'sheet_package_id', 'sheet_stem', 'folder_path', 'qc_pdf_guid'
            )
            if ($cols.Count -gt 0) {
                $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
                $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (50) $select FROM [sheet_index] WHERE folder_path = @folderPath AND document_name = @documentName" -Parameters @{
                    folderPath = $folderPath; documentName = $documentName
                }
                _QDM-IngestIdentityRows -Rows $rows -PackageIds $packageIds -DocumentGuids $documentGuids -DocumentNames $documentNames -SheetStems $sheetStems
            }
        }

        if (_QDM-TestTableExists -TableName 'sheet_packages') {
            $cols = _QDM-SelectExistingColumns -TableName 'sheet_packages' -Requested @(
                'sheet_package_id', 'sheet_stem', 'folder_path', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid'
            )
            if ($cols.Count -gt 0) {
                $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
                $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (20) $select FROM [sheet_packages] WHERE folder_path = @folderPath AND sheet_stem = @sheetStem" -Parameters @{
                    folderPath = $folderPath; sheetStem = $sheetStem
                }
                _QDM-IngestIdentityRows -Rows $rows -PackageIds $packageIds -DocumentGuids $documentGuids -DocumentNames $documentNames -SheetStems $sheetStems
            }
        }

        if (_QDM-TestTableExists -TableName 'sheet_documents') {
            $cols = _QDM-SelectExistingColumns -TableName 'sheet_documents' -Requested @(
                'document_guid', 'document_name', 'sheet_package_id', 'document_role'
            )
            if ($cols.Count -gt 0) {
                $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
                $clauses = @('document_name = @documentName')
                $params = @{ documentName = $documentName }
                if ($packageIds.Count -gt 0) {
                    $inList = (@($packageIds) | ForEach-Object { "'$_'" }) -join ','
                    $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
                }
                $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (50) $select FROM [sheet_documents] WHERE $($clauses -join ' OR ')" -Parameters $params
                _QDM-IngestIdentityRows -Rows $rows -PackageIds $packageIds -DocumentGuids $documentGuids -DocumentNames $documentNames -SheetStems $sheetStems
            }
        }
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($SheetNumber)) {
        $lookupType = 'sheet_number'
        $lookupText = $SheetNumber.Trim()
        $like = _QDM-LikePattern -Text $lookupText
        foreach ($entry in $script:QDM_SheetSearchTables) {
            $table = $entry.table
            if (-not (_QDM-TestTableExists -TableName $table)) { continue }
            $where = _QDM-BuildLikeWhere -TableName $table -Columns $entry.columns -LikeParam $like
            if (-not $where) { continue }
            $cols = _QDM-SelectExistingColumns -TableName $table -Requested @(
                'document_guid', 'document_name', 'sheet_package_id', 'sheet_stem', 'document_role', 'pw_objguid', 'document_id', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid'
            )
            if ($cols.Count -eq 0) { continue }
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            try {
                $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (200) $select FROM [$table] WHERE $($where.clause)" -Parameters $where.params
            } catch {
                $warnings += _QDM-BuildWarning -Message "Search failed on $table`: $($_.Exception.Message)" -Table $table
                continue
            }
            foreach ($row in $rows) {
                foreach ($g in @('document_guid', 'pw_objguid', 'document_id', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid')) {
                    if ($row[$g]) { [void]$documentGuids.Add([string]$row[$g]) }
                }
                if ($row.document_name) { [void]$documentNames.Add([string]$row.document_name) }
                if ($row.sheet_package_id) { [void]$packageIds.Add([string]$row.sheet_package_id) }
                if ($row.sheet_stem) { [void]$sheetStems.Add([string]$row.sheet_stem) }
            }
        }
    }
    else {
        throw 'Provide sheet_number, document_guid, package_id, or document_path.'
    }

    if ($packageIds.Count -eq 0 -and $documentGuids.Count -eq 0 -and $documentNames.Count -eq 0) {
        $warnings += _QDM-BuildWarning -Message 'No matching rows found in telemetry tables.'
    }

    return @{
        lookup_type = $lookupType
        lookup_value = $lookupText
        folder_path = $folderPath
        document_name = $documentName
        sheet_stem = $sheetStem
        raw_document_path = $rawDocumentPath
        sheet_package_ids = @($packageIds)
        document_guids = @($documentGuids)
        document_names = @($documentNames)
        sheet_stems = @($sheetStems)
        warnings = @($warnings)
    }
}

function Search-QCDebugSheet {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $lookup = Resolve-QCDebugLookup @lookupParams
    $likeText = if ($lookup.lookup_type -eq 'document_path' -and $lookup.sheet_stem) { [string]$lookup.sheet_stem } else { [string]$lookup.lookup_value }
    $like = _QDM-LikePattern -Text $likeText
    $grouped = @{}
    $warnings = @()
    if ($lookup.warnings) { $warnings += @($lookup.warnings) }
    $sourceTables = [System.Collections.Generic.List[string]]::new()

    if ($lookup.lookup_type -ne 'document_path') {
    foreach ($entry in $script:QDM_SheetSearchTables) {
        $table = $entry.table
        if (-not (_QDM-TestTableExists -TableName $table)) {
            $warnings += _QDM-BuildWarning -Message "Table dbo.$table is not present; skipped." -Table $table
            continue
        }
        $where = _QDM-BuildLikeWhere -TableName $table -Columns $entry.columns -LikeParam $like
        if (-not $where) {
            $warnings += _QDM-BuildWarning -Message "No searchable text columns on dbo.$table; skipped." -Table $table
            continue
        }
        $cols = _QDM-GetColumns -TableName $table
        $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
        try {
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (100) $select FROM [$table] WHERE $($where.clause) ORDER BY 1 DESC" -Parameters $where.params
        } catch {
            $warnings += _QDM-BuildWarning -Message "Query failed: $($_.Exception.Message)" -Table $table
            continue
        }
        if ($rows.Count -gt 0) {
            $grouped[$table] = $rows
            [void]$sourceTables.Add($table)
        }
    }

    if (_QDM-TestTableExists -TableName 'v_sheet_package_status') {
        $vcols = _QDM-GetColumns -TableName 'v_sheet_package_status'
        $vsearch = @($vcols | Where-Object { $_ -match '(?i)name|stem|package' })
        $where = _QDM-BuildLikeWhere -TableName 'v_sheet_package_status' -Columns $vsearch -LikeParam $like
        if ($where) {
            $select = ($vcols | ForEach-Object { "[$_]" }) -join ', '
            try {
                $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (100) $select FROM [v_sheet_package_status] WHERE $($where.clause)" -Parameters $where.params
                if ($rows.Count -gt 0) {
                    $grouped['v_sheet_package_status'] = $rows
                    [void]$sourceTables.Add('v_sheet_package_status')
                }
            } catch {
                $warnings += _QDM-BuildWarning -Message "View query failed: $($_.Exception.Message)" -Table 'v_sheet_package_status'
            }
        }
    }
    }

    if ($lookup.lookup_type -in @('document_path', 'folder_path')) {
        foreach ($table in @('sheet_index', 'sheet_packages', 'sheet_documents', 'transition_events', 'document_state_history', 'notification_log')) {
            if (-not (_QDM-TestTableExists -TableName $table)) { continue }
            $cols = _QDM-GetColumns -TableName $table
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $clauses = @()
            $params = @{}
            if ($cols -contains 'folder_path') {
                $clauses += 'folder_path = @folderPath'
                $params['folderPath'] = $lookup.folder_path
            }
            if ($lookup.lookup_type -eq 'document_path') {
                if ($table -eq 'sheet_packages' -and ($cols -contains 'sheet_stem')) {
                    $clauses += 'sheet_stem = @sheetStem'
                    $params['sheetStem'] = $lookup.sheet_stem
                }
                elseif ($cols -contains 'document_name') {
                    $clauses += 'document_name = @documentName'
                    $params['documentName'] = $lookup.document_name
                }
            }
            if ($lookup.sheet_package_ids.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
                $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
                $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
            if ($clauses.Count -eq 0) { continue }
            try {
                $orderBy = if ($table -eq 'sheet_index' -or $table -eq 'sheet_packages') { ' ORDER BY last_updated_at DESC' } else { '' }
                $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (100) $select FROM [$table] WHERE $($clauses -join ' AND ')$orderBy" -Parameters $params
            } catch {
                $warnings += _QDM-BuildWarning -Message "Path query failed on $table`: $($_.Exception.Message)" -Table $table
                continue
            }
            if ($rows.Count -gt 0) {
                $grouped[$table] = $rows
                [void]$sourceTables.Add($table)
            }
        }
        if (_QDM-TestTableExists -TableName 'audit_events') {
            $cols = _QDM-GetColumns -TableName 'audit_events'
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            try {
                $auditSql = if ($lookup.lookup_type -eq 'folder_path') {
                    "SELECT TOP (100) $select FROM [audit_events] WHERE resolved_folder = @folderPath ORDER BY captured_at DESC"
                } else {
                    "SELECT TOP (100) $select FROM [audit_events] WHERE resolved_folder = @folderPath AND pw_itemname = @documentName ORDER BY captured_at DESC"
                }
                $auditParams = if ($lookup.lookup_type -eq 'folder_path') {
                    @{ folderPath = $lookup.folder_path }
                } else {
                    @{ folderPath = $lookup.folder_path; documentName = $lookup.document_name }
                }
                $rows = _QDM-RowsFromQuery -Sql $auditSql -Parameters $auditParams
                if ($rows.Count -gt 0) {
                    $grouped['audit_events'] = $rows
                    [void]$sourceTables.Add('audit_events')
                }
            } catch {
                $warnings += _QDM-BuildWarning -Message "Path query failed on audit_events: $($_.Exception.Message)" -Table 'audit_events'
            }
        }
    }

    if ($lookup.lookup_type -eq 'package_id') {
        foreach ($table in @('sheet_packages', 'sheet_documents', 'sheet_index')) {
            if (-not (_QDM-TestTableExists -TableName $table)) { continue }
            if (-not (_QDM-GetColumns -TableName $table | Where-Object { $_ -ieq 'sheet_package_id' })) { continue }
            $cols = _QDM-GetColumns -TableName $table
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (100) $select FROM [$table] WHERE CAST(sheet_package_id AS NVARCHAR(36)) = @pkg" -Parameters @{ pkg = $lookup.lookup_value }
            if ($rows.Count -gt 0 -and -not $grouped.ContainsKey($table)) {
                $grouped[$table] = $rows
                [void]$sourceTables.Add($table)
            }
        }
    }

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        matches_by_table = $grouped
        table_count = $grouped.Keys.Count
    } -Warnings $warnings -SourceTables @($sourceTables) -QueryAssumptions @(
        "Lookup type: $($lookup.lookup_type).",
        'Read-only search across known QC telemetry tables.'
    )
}

function Get-QCDebugSheetIdentity {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $lookup = Resolve-QCDebugLookup @lookupParams
    $warnings = @()
    if ($lookup.warnings) { $warnings += @($lookup.warnings) }
    $sourceTables = [System.Collections.Generic.List[string]]::new()
    $candidates = @{
        sheet_package_ids = [System.Collections.Generic.List[hashtable]]::new()
        document_guids = [System.Collections.Generic.List[hashtable]]::new()
        document_names = [System.Collections.Generic.List[hashtable]]::new()
        sheet_stems = [System.Collections.Generic.List[hashtable]]::new()
        roles = [System.Collections.Generic.List[hashtable]]::new()
    }

    function Add-UniqueCandidate {
        param([string]$Bucket, $Value, [string]$Source, [hashtable]$Extra = @{})
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return }
        $text = [string]$Value
        foreach ($item in $candidates[$Bucket]) {
            if ($item.value -ieq $text) { return }
        }
        $entry = @{ value = $text; source = $Source }
        foreach ($k in $Extra.Keys) { $entry[$k] = $Extra[$k] }
        $candidates[$Bucket].Add($entry)
    }

    $like = _QDM-LikePattern -Text $lookup.lookup_value
    foreach ($table in @('sheet_index', 'sheet_packages', 'sheet_documents')) {
        if (-not (_QDM-TestTableExists -TableName $table)) {
            $warnings += _QDM-BuildWarning -Message "Missing table dbo.$table" -Table $table
            continue
        }
        $cols = _QDM-SelectExistingColumns -TableName $table -Requested @(
            'document_guid', 'document_name', 'sheet_package_id', 'sheet_stem', 'document_role', 'folder_path', 'pw_state_name', 'dgn_guid', 'sheet_pdf_guid', 'qc_pdf_guid'
        )
        if ($cols.Count -eq 0) { continue }
        $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
        $whereParts = @()
        $params = @{}
        if ($lookup.lookup_type -eq 'package_id') {
            $whereParts += 'CAST(sheet_package_id AS NVARCHAR(36)) = @pkg'
            $params['pkg'] = $lookup.lookup_value
        }
        elseif ($lookup.lookup_type -eq 'document_guid') {
            if ($cols -contains 'document_guid') { $whereParts += 'CAST(document_guid AS NVARCHAR(36)) = @docGuid' }
            if ($cols -contains 'dgn_guid') { $whereParts += 'CAST(dgn_guid AS NVARCHAR(36)) = @docGuid' }
            if ($cols -contains 'sheet_pdf_guid') { $whereParts += 'CAST(sheet_pdf_guid AS NVARCHAR(36)) = @docGuid' }
            if ($cols -contains 'qc_pdf_guid') { $whereParts += 'CAST(qc_pdf_guid AS NVARCHAR(36)) = @docGuid' }
            $params['docGuid'] = $lookup.lookup_value
        }
        elseif ($lookup.lookup_type -eq 'document_path') {
            if ($cols -contains 'folder_path') {
                $whereParts += 'folder_path = @folderPath'
                $params['folderPath'] = $lookup.folder_path
            }
            if ($table -eq 'sheet_packages' -and ($cols -contains 'sheet_stem')) {
                $whereParts += 'sheet_stem = @sheetStem'
                $params['sheetStem'] = $lookup.sheet_stem
            }
            elseif ($cols -contains 'document_name') {
                $whereParts += 'document_name = @documentName'
                $params['documentName'] = $lookup.document_name
            }
            if ($lookup.sheet_package_ids.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
                $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
                $whereParts += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
        }
        else {
            $textWhere = _QDM-BuildLikeWhere -TableName $table -Columns @(
                'document_name', 'sheet_stem', 'folder_path', 'document_guid'
            ) -LikeParam $like
            if ($textWhere) {
                $whereParts += "($($textWhere.clause))"
                $params = $textWhere.params
            }
        }
        if ($whereParts.Count -eq 0) { continue }
        $whereJoin = if ($lookup.lookup_type -eq 'document_path') { ' AND ' } else { ' OR ' }
        $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (200) $select FROM [$table] WHERE $($whereParts -join $whereJoin)" -Parameters $params
        [void]$sourceTables.Add($table)
        foreach ($row in $rows) {
            Add-UniqueCandidate -Bucket 'document_guids' -Value $row.document_guid -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'document_guids' -Value $row.dgn_guid -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'document_guids' -Value $row.sheet_pdf_guid -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'document_guids' -Value $row.qc_pdf_guid -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'document_names' -Value $row.document_name -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'sheet_package_ids' -Value $row.sheet_package_id -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'sheet_stems' -Value $row.sheet_stem -Source $table -Extra $row
            Add-UniqueCandidate -Bucket 'roles' -Value $row.document_role -Source $table -Extra $row
        }
    }

    $pkgValues = @($candidates.sheet_package_ids | ForEach-Object { $_.value.ToLowerInvariant() } | Select-Object -Unique)
    if ($pkgValues.Count -gt 1) {
        $warnings += _QDM-BuildWarning -Message 'Multiple distinct sheet_package_id values found; identity may be inconsistent.'
    }
    if ($candidates.document_guids.Count -gt 6) {
        $warnings += _QDM-BuildWarning -Message 'Many document GUID candidates found; sheet may have historical/duplicate rows.'
    }

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        candidates = @{
            sheet_package_ids = @($candidates.sheet_package_ids)
            document_guids = @($candidates.document_guids)
            document_names = @($candidates.document_names)
            sheet_stems = @($candidates.sheet_stems)
            roles = @($candidates.roles)
        }
    } -Warnings $warnings -SourceTables @($sourceTables) -QueryAssumptions @(
        'Aggregates identity hints from available package/document tables only.'
    )
}

function Get-QCDebugSheetPackageMembers {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $identity = Get-QCDebugSheetIdentity @lookupParams
    $lookup = $identity.data.lookup
    $warnings = @()
    if ($identity.warnings) { $warnings += @($identity.warnings) }
    $packageIds = @($identity.data.candidates.sheet_package_ids | ForEach-Object { $_.value })
    $members = @{
        by_role = @{}
        sheet_index_rows = @()
        sheet_packages_rows = @()
        sheet_documents_rows = @()
    }
    $sourceTables = [System.Collections.Generic.List[string]]::new()

    if (_QDM-TestTableExists -TableName 'sheet_packages') {
        [void]$sourceTables.Add('sheet_packages')
        $cols = _QDM-SelectExistingColumns -TableName 'sheet_packages' -Requested @(
            'sheet_package_id', 'sheet_stem', 'folder_path', 'pw_state_name', 'dgn_guid', 'dgn_name', 'sheet_pdf_guid', 'sheet_pdf_name', 'qc_pdf_guid', 'qc_pdf_name', 'designer_email', 'reviewer_email', 'checker_email'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $clauses = @()
            $params = @{}
            if ($lookup.lookup_type -eq 'sheet_number') {
                if ($cols -contains 'sheet_stem') {
                    $clauses += 'sheet_stem LIKE @like'
                    $params['like'] = _QDM-LikePattern -Text $lookup.lookup_value
                }
            }
            if ($lookup.lookup_type -eq 'document_path') {
                if ($cols -contains 'folder_path') {
                    $clauses += 'folder_path = @folderPath'
                    $params['folderPath'] = $lookup.folder_path
                }
                if ($cols -contains 'sheet_stem') {
                    $clauses += 'sheet_stem = @sheetStem'
                    $params['sheetStem'] = $lookup.sheet_stem
                }
            }
            if ($packageIds.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
                $inList = ($packageIds | ForEach-Object { "'$_'" }) -join ','
                $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
            if ($lookup.lookup_type -eq 'package_id') {
                $clauses += 'CAST(sheet_package_id AS NVARCHAR(36)) = @pkg'
                $params['pkg'] = $lookup.lookup_value
            }
            if ($clauses.Count -gt 0) {
                $clauseJoin = if ($lookup.lookup_type -eq 'document_path') { ' AND ' } else { ' OR ' }
                $members.sheet_packages_rows = [object[]]@(_QDM-RowsFromQuery -Sql "SELECT $select FROM [sheet_packages] WHERE $($clauses -join $clauseJoin)" -Parameters $params)
            }
        }
    }

    if (_QDM-TestTableExists -TableName 'sheet_documents') {
        [void]$sourceTables.Add('sheet_documents')
        $cols = _QDM-SelectExistingColumns -TableName 'sheet_documents' -Requested @(
            'document_guid', 'document_name', 'document_role', 'pw_state_name', 'sheet_package_id', 'extension'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $clauses = @()
            $params = @{}
            if ($lookup.lookup_type -eq 'sheet_number') {
                if ($cols -contains 'document_name') {
                    $clauses += 'document_name LIKE @like'
                    $params['like'] = _QDM-LikePattern -Text $lookup.lookup_value
                }
            }
            if ($lookup.lookup_type -eq 'document_path' -and ($cols -contains 'document_name')) {
                $clauses += 'document_name = @documentName'
                $params['documentName'] = $lookup.document_name
            }
            if ($packageIds.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
                $inList = ($packageIds | ForEach-Object { "'$_'" }) -join ','
                $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
            if ($lookup.lookup_type -eq 'document_guid') {
                $clauses += 'CAST(document_guid AS NVARCHAR(36)) = @docGuid'
                $params['docGuid'] = $lookup.lookup_value
            }
            if ($clauses.Count -gt 0) {
                $members.sheet_documents_rows = [object[]]@(_QDM-RowsFromQuery -Sql "SELECT $select FROM [sheet_documents] WHERE $($clauses -join ' OR ')" -Parameters $params)
                foreach ($row in $members.sheet_documents_rows) {
                    $role = [string]($row.document_role)
                    if ([string]::IsNullOrWhiteSpace($role)) { $role = 'unknown' }
                    $members.by_role[$role] = $row
                }
            }
        }
    }

    if (_QDM-TestTableExists -TableName 'sheet_index') {
        [void]$sourceTables.Add('sheet_index')
        $cols = _QDM-SelectExistingColumns -TableName 'sheet_index' -Requested @(
            'document_guid', 'document_name', 'extension', 'pw_state_name', 'sheet_package_id', 'folder_path', 'qc_pdf_guid', 'qc_process_type', 'qc_review_type', 'last_updated_at'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $clauses = @()
            $params = @{}
            if ($lookup.lookup_type -eq 'sheet_number') {
                $clauses += '(document_name LIKE @like OR folder_path LIKE @like)'
                $params['like'] = _QDM-LikePattern -Text $lookup.lookup_value
            }
            if ($lookup.lookup_type -eq 'document_path') {
                $clauses += '(folder_path = @folderPath AND document_name = @documentName)'
                $params['folderPath'] = $lookup.folder_path
                $params['documentName'] = $lookup.document_name
            }
            if ($lookup.lookup_type -eq 'document_guid') {
                $clauses += '(CAST(document_guid AS NVARCHAR(36)) = @docGuid OR CAST(qc_pdf_guid AS NVARCHAR(36)) = @docGuid)'
                $params['docGuid'] = $lookup.lookup_value
            }
            if ($packageIds.Count -gt 0) {
                $inList = ($packageIds | ForEach-Object { "'$_'" }) -join ','
                $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
            if ($clauses.Count -gt 0) {
                $order = if ($cols -contains 'last_updated_at') { ' ORDER BY last_updated_at DESC' } else { '' }
                $members.sheet_index_rows = [object[]]@(_QDM-RowsFromQuery -Sql "SELECT TOP (50) $select FROM [sheet_index] WHERE $($clauses -join ' OR ')$order" -Parameters $params)
            }
        }
    }

    $pkgRowsLocal = @()
    if ($members.sheet_packages_rows) { $pkgRowsLocal = @($members.sheet_packages_rows) }
    $pkg = @{}
    if ($pkgRowsLocal.Count -gt 0 -and $null -ne $pkgRowsLocal[0]) { $pkg = $pkgRowsLocal[0] }
    foreach ($roleInfo in @(
        @{ role = 'dgn'; guid = 'dgn_guid' }
        @{ role = 'sheet_pdf'; guid = 'sheet_pdf_guid' }
        @{ role = 'qc_pdf'; guid = 'qc_pdf_guid' }
    )) {
        $pkgGuid = [string]($pkg[$roleInfo.guid])
        $docRow = $members.by_role[$roleInfo.role]
        $docGuid = if ($docRow) { [string]$docRow.document_guid } else { '' }
        if ($pkgGuid -and $docGuid -and ($pkgGuid.ToLowerInvariant() -ne $docGuid.ToLowerInvariant())) {
            $warnings += _QDM-BuildWarning -Message "GUID mismatch for role $($roleInfo.role): sheet_packages.$($roleInfo.guid)=$pkgGuid vs sheet_documents=$docGuid" -Table 'sheet_packages' -Column $roleInfo.guid
        }
    }

    $packageIdsForLanes = @($packageIds)
    if ($pkg.sheet_package_id) { $packageIdsForLanes += [string]$pkg.sheet_package_id }
    $packageIdsForLanes = @($packageIdsForLanes | Select-Object -Unique)
    $members['sheet_package_qc_pdf_rows'] = @(_QDM-LoadSheetPackageQcPdfRows -PackageIds $packageIdsForLanes)
    $members['qc_process_type'] = _QDM-BuildQcProcessTypeDiagnostics -Lookup $lookup -Members $members -ProjectWiseAvailable $false

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        members = $members
    } -Warnings $warnings -SourceTables @($sourceTables) -QueryAssumptions @(
        'Compares package registry, role table, and sheet_index when present.'
    )
}

function _QDM-NormalizeQcProcessTypeValue {
    param([string]$RawProcessType)
    if ([string]::IsNullOrWhiteSpace($RawProcessType)) { return $null }
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        return Normalize-QCProcessType -ProcessType ([string]$RawProcessType) -AllowNullOnEmpty
    }
    $text = ([string]$RawProcessType).Trim().ToLowerInvariant()
    switch ($text) {
        'prod' { return 'production' }
        'chk' { return 'check' }
        'rev' { return 'review' }
        default { return $text }
    }
}

function _QDM-InferQcProcessTypeFromDocumentName {
    param([string]$DocumentName)
    if ([string]::IsNullOrWhiteSpace($DocumentName)) { return $null }
    if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
        return Get-PWQcPdfLaneFromDocumentName -DocumentName ([string]$DocumentName)
    }
    $name = [System.IO.Path]::GetFileName([string]$DocumentName)
    if ($name -match '(?i)-prod\.pdf$') { return 'production' }
    if ($name -match '(?i)-chk\.pdf$') { return 'check' }
    if ($name -match '(?i)-rev\.pdf$') { return 'review' }
    return $null
}

function _QDM-ResolveMemberFolderPath {
    param(
        [hashtable]$Lookup,
        [hashtable]$PkgRow,
        [array]$IndexRows
    )
    if ($Lookup -and $Lookup.folder_path -and -not [string]::IsNullOrWhiteSpace([string]$Lookup.folder_path)) {
        return [string]$Lookup.folder_path
    }
    if ($PkgRow -and $PkgRow.folder_path -and -not [string]::IsNullOrWhiteSpace([string]$PkgRow.folder_path)) {
        return [string]$PkgRow.folder_path
    }
    foreach ($row in @($IndexRows)) {
        if ($row.folder_path -and -not [string]::IsNullOrWhiteSpace([string]$row.folder_path)) {
            return [string]$row.folder_path
        }
    }
    return ''
}

function _QDM-LoadSheetPackageQcPdfRows {
    param([string[]]$PackageIds)
    if ($PackageIds.Count -eq 0) { return @() }
    if (-not (_QDM-TestTableExists -TableName 'sheet_package_qc_pdfs')) { return @() }
    $cols = _QDM-SelectExistingColumns -TableName 'sheet_package_qc_pdfs' -Requested @(
        'sheet_package_id', 'qc_process_type', 'document_guid', 'document_name', 'is_active', 'created_at', 'updated_at'
    )
    if ($cols.Count -eq 0) { return @() }
    $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
    $inList = (@($PackageIds | Where-Object { $_ }) | ForEach-Object { "'$_'" }) -join ','
    if ([string]::IsNullOrWhiteSpace($inList)) { return @() }
    return [object[]]@(_QDM-RowsFromQuery -Sql "SELECT $select FROM [sheet_package_qc_pdfs] WHERE CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList) ORDER BY qc_process_type" -Parameters @{})
}

function _QDM-ReadProjectWiseProcessTypesForDocuments {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [array]$Documents,
        [System.Collections.Generic.List[object]]$Warnings = $null
    )
    $processTypeMap = @{}
    if ($Documents.Count -eq 0) { return $processTypeMap }
    $pw = $Config.projectWise
    $ds = [string]$pw.datasourceName
    $credPath = [string]$pw.credentialPath
    if ([string]::IsNullOrWhiteSpace($ds) -or [string]::IsNullOrWhiteSpace($credPath)) {
        if ($Warnings) { $Warnings.Add((_QDM-BuildWarning -Message 'projectWise.datasourceName or credentialPath missing; skipping live QC_Process_Type reads.')) }
        return $processTypeMap
    }
    if (-not (Get-Command -Name 'Invoke-PWAuthenticatedCommand' -ErrorAction SilentlyContinue)) {
        if ($Warnings) { $Warnings.Add((_QDM-BuildWarning -Message 'Invoke-PWAuthenticatedCommand unavailable; skipping live QC_Process_Type reads.')) }
        return $processTypeMap
    }
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        if ($Warnings) { $Warnings.Add((_QDM-BuildWarning -Message 'Folder path unknown; skipping live QC_Process_Type reads.')) }
        return $processTypeMap
    }

    $docInputs = @($Documents | ForEach-Object {
        @{
            guid = [string]$_.document_guid
            name = [string]$_.document_name
        }
    } | Where-Object { $_.guid -and $_.name })
    if ($docInputs.Count -eq 0) { return $processTypeMap }

    $modulesRoot = Split-Path -Parent $PSScriptRoot
    $cfgLocal = $Config
    $folderLocal = $FolderPath
    try {
        $pwResult = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -KeepSession -ScriptBlock {
            Import-Module (Join-Path $modulesRoot 'ProjectWise/PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue | Out-Null
            $processTypes = @{}
            foreach ($doc in $docInputs) {
                $guid = [string]$doc.guid
                $name = [string]$doc.name
                if ([string]::IsNullOrWhiteSpace($guid) -or [string]::IsNullOrWhiteSpace($name)) { continue }
                $raw = ''
                if (Get-Command -Name 'Get-PWQcPrependProcessIntentFromSourcePdf' -ErrorAction SilentlyContinue) {
                    $read = Get-PWQcPrependProcessIntentFromSourcePdf -FolderPath $folderLocal -SourceDocumentName $name -Config $cfgLocal
                    if ($read.found -and $read.qcProcessType) { $raw = [string]$read.qcProcessType }
                } elseif (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
                    $read = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $folderLocal -SourceDocumentName $name -Config $cfgLocal
                    if ($read.found -and $read.qcProcessType) { $raw = [string]$read.qcProcessType }
                }
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $processTypes[$guid.ToLowerInvariant()] = $raw
                }
            }
            return $processTypes
        }
        if ($pwResult -is [hashtable]) {
            foreach ($k in $pwResult.Keys) { $processTypeMap[[string]$k] = [string]$pwResult[$k] }
        }
    } catch {
        if ($Warnings) { $Warnings.Add((_QDM-BuildWarning -Message "ProjectWise QC_Process_Type read failed: $($_.Exception.Message)")) }
    }
    return $processTypeMap
}

function _QDM-BuildQcProcessTypeDiagnostics {
    param(
        [hashtable]$Lookup,
        [hashtable]$Members,
        [hashtable]$PwProcessTypeByGuid = @{},
        [bool]$ProjectWiseAvailable = $false
    )

    $indexRows = @($Members.sheet_index_rows)
    $docRows = @($Members.sheet_documents_rows)
    $pkgRows = @($Members.sheet_packages_rows)
    $pkg = if ($pkgRows.Count -gt 0) { $pkgRows[0] } else { @{} }
    $folderPath = _QDM-ResolveMemberFolderPath -Lookup $Lookup -PkgRow $pkg -IndexRows $indexRows

    $packageIds = @($docRows | ForEach-Object { [string]$_.sheet_package_id } | Where-Object { $_ })
    if ($pkg.sheet_package_id) { $packageIds += [string]$pkg.sheet_package_id }
    $packageIds = @($packageIds | Select-Object -Unique)
    $laneRegistryRows = @(_QDM-LoadSheetPackageQcPdfRows -PackageIds $packageIds)

    $indexByGuid = @{}
    $indexByName = @{}
    foreach ($row in $indexRows) {
        $norm = _QDM-NormalizeQcProcessTypeValue -RawProcessType ([string]$row.qc_process_type)
        if ($row.document_guid) { $indexByGuid[[string]$row.document_guid.ToLowerInvariant()] = $norm }
        if ($row.document_name) { $indexByName[[string]$row.document_name.ToLowerInvariant()] = $norm }
    }

    $laneByGuid = @{}
    foreach ($row in $laneRegistryRows) {
        $norm = _QDM-NormalizeQcProcessTypeValue -RawProcessType ([string]$row.qc_process_type)
        if ($row.document_guid) { $laneByGuid[[string]$row.document_guid.ToLowerInvariant()] = $norm }
    }

    $documents = [System.Collections.Generic.List[hashtable]]::new()
    $seenGuids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $docRows) {
        if (-not $row.document_guid) { continue }
        $guid = [string]$row.document_guid
        if ($seenGuids.Contains($guid)) { continue }
        [void]$seenGuids.Add($guid)
        $documents.Add(@{
            document_guid = $guid
            document_name = [string]$row.document_name
            document_role = [string]$row.document_role
            source = 'sheet_documents'
        })
    }
    foreach ($row in $indexRows) {
        if (-not $row.document_guid) { continue }
        $guid = [string]$row.document_guid
        if ($seenGuids.Contains($guid)) { continue }
        [void]$seenGuids.Add($guid)
        $documents.Add(@{
            document_guid = $guid
            document_name = [string]$row.document_name
            document_role = ''
            source = 'sheet_index'
        })
    }

    $checks = [System.Collections.Generic.List[hashtable]]::new()
    $rows = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($doc in $documents) {
        $guidKey = [string]$doc.document_guid.ToLowerInvariant()
        $nameKey = if ($doc.document_name) { [string]$doc.document_name.ToLowerInvariant() } else { '' }
        $filenameInferred = _QDM-InferQcProcessTypeFromDocumentName -DocumentName ([string]$doc.document_name)
        $dbIndex = if ($indexByGuid.ContainsKey($guidKey)) { $indexByGuid[$guidKey] } elseif ($nameKey -and $indexByName.ContainsKey($nameKey)) { $indexByName[$nameKey] } else { $null }
        $dbLane = if ($laneByGuid.ContainsKey($guidKey)) { $laneByGuid[$guidKey] } else { $null }
        $pwRaw = if ($PwProcessTypeByGuid.ContainsKey($guidKey)) { [string]$PwProcessTypeByGuid[$guidKey] } else { '' }
        $pwNorm = _QDM-NormalizeQcProcessTypeValue -RawProcessType $pwRaw
        $isLanePdf = $null -ne $filenameInferred

        $rows.Add(@{
            document_guid = [string]$doc.document_guid
            document_name = [string]$doc.document_name
            document_role = [string]$doc.document_role
            source = [string]$doc.source
            is_lane_pdf = $isLanePdf
            filename_inferred_process_type = $filenameInferred
            database_sheet_index_process_type = $dbIndex
            database_lane_registry_process_type = $dbLane
            projectwise_process_type = if ($pwRaw) { $pwRaw } else { $null }
            projectwise_process_type_normalized = $pwNorm
        })

        if ($isLanePdf) {
            if ($dbLane -and $dbLane -ne $filenameInferred) {
                $checks.Add(@{
                    code = 'lane_registry_vs_filename'
                    passed = $false
                    document_guid = [string]$doc.document_guid
                    document_name = [string]$doc.document_name
                    message = "sheet_package_qc_pdfs has '$dbLane' but filename implies '$filenameInferred'."
                    expected = $filenameInferred
                    actual = $dbLane
                })
            }
            if ($ProjectWiseAvailable -and $pwNorm -and $pwNorm -ne $filenameInferred) {
                $checks.Add(@{
                    code = 'projectwise_vs_lane_filename'
                    passed = $false
                    document_guid = [string]$doc.document_guid
                    document_name = [string]$doc.document_name
                    message = "ProjectWise QC_Process_Type '$pwRaw' disagrees with lane filename suffix ('$filenameInferred')."
                    expected = $filenameInferred
                    actual = $pwNorm
                })
            }
        } else {
            if ($dbIndex -and $pwNorm -and $dbIndex -ne $pwNorm) {
                $checks.Add(@{
                    code = 'sheet_index_vs_projectwise_process_type'
                    passed = $false
                    document_guid = [string]$doc.document_guid
                    document_name = [string]$doc.document_name
                    message = "sheet_index qc_process_type '$dbIndex' differs from ProjectWise '$pwRaw'."
                    expected = $pwNorm
                    actual = $dbIndex
                })
            }
        }
    }

    foreach ($laneRow in $laneRegistryRows) {
        $proc = _QDM-NormalizeQcProcessTypeValue -RawProcessType ([string]$laneRow.qc_process_type)
        $docName = [string]$laneRow.document_name
        $suffixExpected = switch ($proc) {
            'production' { 'prod' }
            'check' { 'chk' }
            'review' { 'rev' }
            default { $null }
        }
        if ($suffixExpected -and $docName -and ($docName -notmatch "(?i)-$suffixExpected\.pdf$")) {
            $checks.Add(@{
                code = 'lane_registry_name_suffix_mismatch'
                passed = $false
                document_guid = [string]$laneRow.document_guid
                document_name = $docName
                message = "Lane registry qc_process_type '$proc' expects *-$suffixExpected.pdf but document_name is '$docName'."
                expected = "*-$suffixExpected.pdf"
                actual = $docName
            })
        }
    }

    $failed = @($checks | Where-Object { -not $_.passed })
    return @{
        folder_path = $folderPath
        documents = @($rows)
        lane_registry = @($laneRegistryRows)
        checks = @($checks)
        checks_passed = [Math]::Max(0, $checks.Count - $failed.Count)
        checks_failed = $failed.Count
        projectwise_available = [bool]$ProjectWiseAvailable
        resolution_policy = 'Lane PDFs (*-prod/*-chk/*-rev): canonical process type comes from the filename suffix. Stem/DGN: compare sheet_index.qc_process_type with live ProjectWise QC_Process_Type.'
    }
}

function Get-QCDebugQcProcessTypeDiagnostics {
    <#
    .SYNOPSIS
    Compare qc_process_type across lane filenames, sheet_index, lane registry, and ProjectWise.
    #>
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $membersResult = Get-QCDebugSheetPackageMembers @lookupParams
    $lookup = $membersResult.data.lookup
    $members = $membersResult.data.members
    $warnings = [System.Collections.Generic.List[object]]::new()
    if ($membersResult.warnings) { foreach ($w in @($membersResult.warnings)) { $warnings.Add($w) } }

    $cfg = _QDM-Config
    $pkgRowsLocal = @($members.sheet_packages_rows)
    $pkgRow = if ($pkgRowsLocal.Count -gt 0) { $pkgRowsLocal[0] } else { @{} }
    $folderPath = _QDM-ResolveMemberFolderPath -Lookup $lookup -PkgRow $pkgRow -IndexRows @($members.sheet_index_rows)
    $docInputs = @($members.sheet_documents_rows | ForEach-Object {
        @{ document_guid = [string]$_.document_guid; document_name = [string]$_.document_name }
    } | Where-Object { $_.document_guid -and $_.document_name })
    $pwProcessTypes = _QDM-ReadProjectWiseProcessTypesForDocuments -Config $cfg -FolderPath $folderPath -Documents $docInputs -Warnings $warnings

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        qc_process_type = (_QDM-BuildQcProcessTypeDiagnostics -Lookup $lookup -Members $members -PwProcessTypeByGuid $pwProcessTypes -ProjectWiseAvailable ($pwProcessTypes.Count -gt 0))
    } -Warnings @($warnings) -SourceTables @('sheet_index', 'sheet_documents', 'sheet_package_qc_pdfs') -QueryAssumptions @(
        'Lane PDF process type is validated against *-prod/*-chk/*-rev filename suffixes.',
        'Stem/DGN process type compares sheet_index with live ProjectWise QC_Process_Type when PW is reachable.'
    )
}

function _QDM-TimelineRow {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][hashtable]$Row,
        [Parameter(Mandatory)][hashtable]$Mapping
    )
    $details = [System.Collections.Generic.List[string]]::new()
    foreach ($col in @($Mapping.detail_cols)) {
        if (-not $col) { continue }
        $val = $Row[$col]
        if ($null -ne $val -and -not [string]::IsNullOrWhiteSpace([string]$val)) {
            $details.Add("$col=$val")
        }
    }
    function Pick([string]$Field) {
        if ([string]::IsNullOrWhiteSpace($Field)) { return $null }
        return $Row[$Field]
    }
    return @{
        event_time = Pick $Mapping.time
        source = $Source
        source_id = Pick $Mapping.id
        document_name = Pick $Mapping.document_name
        document_guid = Pick $Mapping.document_guid
        sheet_package_id = Pick $Mapping.sheet_package_id
        event_type = Pick $Mapping.event_type
        state_from = Pick $Mapping.state_from
        state_to = Pick $Mapping.state_to
        status = Pick $Mapping.status
        actor = Pick $Mapping.actor
        details = if ($details.Count -gt 0) { ($details -join '; ') } else { $null }
    }
}

function Get-QCDebugSheetTimeline {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = '',
        [int]$Limit = 200
    )

    $top = _QDM-SafeTopLimit -Limit $Limit -MaxLimit 500
    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $lookup = Resolve-QCDebugLookup @lookupParams
    $warnings = @()
    if ($lookup.warnings) { $warnings += @($lookup.warnings) }
    $events = [System.Collections.Generic.List[hashtable]]::new()
    $sourceTables = [System.Collections.Generic.List[string]]::new()
    $likeText = if ($lookup.lookup_type -eq 'document_path' -and $lookup.sheet_stem) { [string]$lookup.sheet_stem } else { [string]$lookup.lookup_value }
    $like = _QDM-LikePattern -Text $likeText

    foreach ($spec in $script:QDM_TimelineSources) {
        $table = [string]$spec.table
        if (-not (_QDM-TestTableExists -TableName $table)) {
            $warnings += _QDM-BuildWarning -Message "Timeline source dbo.$table not present; skipped." -Table $table
            continue
        }
        $timeCol = [string]$spec.time
        if (-not ((_QDM-GetColumns -TableName $table) -contains $timeCol)) {
            $warnings += _QDM-BuildWarning -Message "Required time column $timeCol missing on dbo.$table; skipped." -Table $table -Column $timeCol
            continue
        }
        $needed = @()
        foreach ($field in @('id', 'time', 'document_name', 'document_guid', 'sheet_package_id', 'event_type', 'state_from', 'state_to', 'status', 'actor')) {
            $col = $spec[$field]
            if ($col) { $needed += [string]$col }
        }
        $needed += @($spec.where_cols)
        $needed += @($spec.detail_cols)
        $selectCols = @(_QDM-SelectExistingColumns -TableName $table -Requested $needed | Select-Object -Unique)
        if ($selectCols.Count -eq 0) {
            $warnings += _QDM-BuildWarning -Message "No usable columns on dbo.$table; skipped." -Table $table
            continue
        }
        $whereParts = @()
        $params = @{}
        if ($lookup.lookup_type -eq 'sheet_number') {
            $textWhere = _QDM-BuildLikeWhere -TableName $table -Columns @($spec.where_cols) -LikeParam $like
            if ($textWhere) {
                $whereParts += "($($textWhere.clause))"
                $params = $textWhere.params
            }
        }
        elseif ($lookup.lookup_type -eq 'document_path') {
            $tableCols = _QDM-GetColumns -TableName $table
            $pathParts = @()
            if ($tableCols -contains 'folder_path') {
                $pathParts += 'folder_path = @folderPath'
                $params['folderPath'] = $lookup.folder_path
            }
            if ($table -eq 'audit_events' -and ($tableCols -contains 'resolved_folder')) {
                $pathParts += 'resolved_folder = @folderPath'
                $params['folderPath'] = $lookup.folder_path
            }
            if ($table -eq 'audit_events' -and ($tableCols -contains 'pw_itemname')) {
                $pathParts += 'pw_itemname = @documentName'
                $params['documentName'] = $lookup.document_name
            }
            elseif ($tableCols -contains 'document_name') {
                $pathParts += 'document_name = @documentName'
                $params['documentName'] = $lookup.document_name
            }
            if ($pathParts.Count -gt 0) {
                $whereParts += '(' + ($pathParts -join ' AND ') + ')'
            }
        }
        $guidCol = [string]$spec.document_guid
        $tableCols = _QDM-GetColumns -TableName $table
        if ($lookup.document_guids.Count -gt 0 -and $guidCol -and ($tableCols -contains $guidCol)) {
            $inList = (@($lookup.document_guids) | ForEach-Object { "'$($_.ToLowerInvariant())'" }) -join ','
            $whereParts += "LOWER(CAST([$guidCol] AS NVARCHAR(36))) IN ($inList)"
        }
        if ($lookup.sheet_package_ids.Count -gt 0 -and ($tableCols -contains 'sheet_package_id')) {
            $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
            $whereParts += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
        }
        if ($whereParts.Count -eq 0) {
            $warnings += _QDM-BuildWarning -Message "No WHERE strategy for dbo.$table; skipped." -Table $table
            continue
        }
        $select = ($selectCols | ForEach-Object { "[$_]" }) -join ', '
        $order = if ($selectCols -contains $timeCol) { " ORDER BY [$timeCol] DESC" } else { '' }
        try {
            $rows = _QDM-RowsFromQuery -Sql "SELECT TOP ($top) $select FROM [$table] WHERE $($whereParts -join ' OR ')$order" -Parameters $params
        } catch {
            $warnings += _QDM-BuildWarning -Message "Timeline query failed on dbo.$table`: $($_.Exception.Message)" -Table $table
            continue
        }
        [void]$sourceTables.Add($table)
        foreach ($row in $rows) {
            $events.Add((_QDM-TimelineRow -Source $spec.source -Row $row -Mapping $spec))
        }
    }

    $sorted = @($events | Sort-Object { $_.event_time } -Descending)
    if ($sorted.Count -gt $top) { $sorted = $sorted[0..($top - 1)] }

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        events = $sorted
        event_count = $sorted.Count
    } -Warnings $warnings -SourceTables @($sourceTables) -QueryAssumptions @(
        "Returns at most $top events after merge.",
        'Skipped sources are reported in warnings.'
    )
}

function Get-QCDebugNotificationDiagnostics {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = '',
        [int]$Limit = 100
    )

    $top = _QDM-SafeTopLimit -Limit $Limit -MaxLimit 200
    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $lookup = Resolve-QCDebugLookup @lookupParams
    $warnings = @()
    if ($lookup.warnings) { $warnings += @($lookup.warnings) }
    $sourceTables = [System.Collections.Generic.List[string]]::new()
    $likeText = if ($lookup.lookup_type -eq 'document_path' -and $lookup.sheet_stem) { [string]$lookup.sheet_stem } else { [string]$lookup.lookup_value }
    $like = _QDM-LikePattern -Text $likeText
    $notifications = @()
    $jobs = @()
    $transitions = @()
    $recipients = @()

    if (_QDM-TestTableExists -TableName 'notification_log') {
        [void]$sourceTables.Add('notification_log')
        $cols = _QDM-SelectExistingColumns -TableName 'notification_log' -Requested @(
            'id', 'sent_at', 'event_type', 'document_guid', 'document_name', 'recipients', 'subject', 'success', 'error_message', 'dedupe_key', 'transition_id', 'sheet_package_id', 'folder_path'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $clauses = @()
            $params = @{}
            if ($lookup.lookup_type -eq 'sheet_number') {
                $textWhere = _QDM-BuildLikeWhere -TableName 'notification_log' -Columns @('document_name', 'subject', 'folder_path') -LikeParam $like
                if ($textWhere) {
                    $clauses += "($($textWhere.clause))"
                    $params = $textWhere.params
                }
            }
            elseif ($lookup.lookup_type -eq 'document_path') {
                if ($cols -contains 'folder_path') {
                    $clauses += '(folder_path = @folderPath AND document_name = @documentName)'
                    $params['folderPath'] = $lookup.folder_path
                    $params['documentName'] = $lookup.document_name
                }
            }
            if ($lookup.document_guids.Count -gt 0 -and ($cols -contains 'document_guid')) {
                $inList = (@($lookup.document_guids) | ForEach-Object { "'$($_.ToLowerInvariant())'" }) -join ','
                $clauses += "LOWER(CAST(document_guid AS NVARCHAR(36))) IN ($inList)"
            }
            if ($lookup.sheet_package_ids.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
                $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
                $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
            if ($clauses.Count -gt 0) {
                $order = if ($cols -contains 'sent_at') { ' ORDER BY sent_at DESC' } else { '' }
                $notifications = _QDM-RowsFromQuery -Sql "SELECT TOP ($top) $select FROM [notification_log] WHERE $($clauses -join ' OR ')$order" -Parameters $params
            }
        }
    }

    if (_QDM-TestTableExists -TableName 'processing_jobs') {
        [void]$sourceTables.Add('processing_jobs')
        $cols = _QDM-SelectExistingColumns -TableName 'processing_jobs' -Requested @(
            'id', 'job_id', 'job_type', 'status', 'created_at', 'completed_at', 'source_path', 'dedupe_key', 'error_code', 'error_message', 'sheet_package_id', 'source_folder', 'worker_machine_name', 'worker_pid'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $sheetFilters = @()
            $params = @{}
            if ($lookup.lookup_type -eq 'sheet_number') {
                $textWhere = _QDM-BuildLikeWhere -TableName 'processing_jobs' -Columns @('source_path', 'source_folder', 'dedupe_key', 'job_id') -LikeParam $like
                if ($textWhere) {
                    $sheetFilters += "($($textWhere.clause))"
                    $params = $textWhere.params
                }
            }
            elseif ($lookup.lookup_type -eq 'document_path') {
                if ($cols -contains 'source_folder') {
                    $sheetFilters += 'source_folder = @folderPath'
                    $params['folderPath'] = $lookup.folder_path
                }
                if ($cols -contains 'source_path') {
                    $sheetFilters += 'source_path LIKE @sourcePathLike'
                    $params['sourcePathLike'] = ('%' + $lookup.document_name)
                }
            }
            if ($lookup.sheet_package_ids.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
                $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
                $sheetFilters += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
            }
            if ($sheetFilters.Count -gt 0 -and ($cols -contains 'job_type')) {
                $order = if ($cols -contains 'created_at') { ' ORDER BY created_at DESC' } else { '' }
                $jobs = _QDM-RowsFromQuery -Sql "SELECT TOP ($top) $select FROM [processing_jobs] WHERE job_type = @jobType AND ($($sheetFilters -join ' OR '))$order" -Parameters (@{ jobType = 'QC_NOTIFICATION' } + $params)
            }
        }
    }

    if (_QDM-TestTableExists -TableName 'transition_events') {
        [void]$sourceTables.Add('transition_events')
        $cols = _QDM-SelectExistingColumns -TableName 'transition_events' -Requested @(
            'id', 'detected_at', 'document_name', 'document_guid', 'from_value', 'to_value', 'transition_type', 'notification_sent', 'notification_id', 'trigger_audit_id', 'folder_path'
        )
        if ($cols.Count -gt 0) {
            $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
            $clauses = @()
            $params = @{}
            if ($lookup.lookup_type -eq 'sheet_number') {
                $textWhere = _QDM-BuildLikeWhere -TableName 'transition_events' -Columns @('document_name', 'folder_path') -LikeParam $like
                if ($textWhere) {
                    $clauses += "($($textWhere.clause))"
                    $params = $textWhere.params
                }
            }
            elseif ($lookup.lookup_type -eq 'document_path' -and ($cols -contains 'folder_path') -and ($cols -contains 'document_name')) {
                $clauses += '(folder_path = @folderPath AND document_name = @documentName)'
                $params['folderPath'] = $lookup.folder_path
                $params['documentName'] = $lookup.document_name
            }
            if ($lookup.document_guids.Count -gt 0 -and ($cols -contains 'document_guid')) {
                $inList = (@($lookup.document_guids) | ForEach-Object { "'$($_.ToLowerInvariant())'" }) -join ','
                $clauses += "LOWER(CAST(document_guid AS NVARCHAR(36))) IN ($inList)"
            }
            if ($clauses.Count -gt 0) {
                $order = if ($cols -contains 'detected_at') { ' ORDER BY detected_at DESC' } else { '' }
                $transitions = _QDM-RowsFromQuery -Sql "SELECT TOP ($top) $select FROM [transition_events] WHERE $($clauses -join ' OR ')$order" -Parameters $params
            }
        }
    }

    foreach ($table in @('sheet_index', 'sheet_packages')) {
        if (-not (_QDM-TestTableExists -TableName $table)) { continue }
        $cols = _QDM-SelectExistingColumns -TableName $table -Requested @('designer_email', 'reviewer_email', 'checker_email', 'sheet_stem', 'document_name', 'sheet_package_id')
        if ($cols.Count -eq 0) { continue }
        $select = ($cols | ForEach-Object { "[$_]" }) -join ', '
        $clauses = @()
        $params = @{}
        if ($lookup.lookup_type -eq 'sheet_number') {
            $textWhere = _QDM-BuildLikeWhere -TableName $table -Columns @('document_name', 'sheet_stem') -LikeParam $like
            if ($textWhere) {
                $clauses += "($($textWhere.clause))"
                $params = $textWhere.params
            }
        }
        elseif ($lookup.lookup_type -eq 'document_path') {
            if ($table -eq 'sheet_index' -and ($cols -contains 'folder_path') -and ($cols -contains 'document_name')) {
                $clauses += '(folder_path = @folderPath AND document_name = @documentName)'
                $params['folderPath'] = $lookup.folder_path
                $params['documentName'] = $lookup.document_name
            }
            elseif ($table -eq 'sheet_packages' -and ($cols -contains 'folder_path') -and ($cols -contains 'sheet_stem')) {
                $clauses += '(folder_path = @folderPath AND sheet_stem = @sheetStem)'
                $params['folderPath'] = $lookup.folder_path
                $params['sheetStem'] = $lookup.sheet_stem
            }
        }
        if ($lookup.sheet_package_ids.Count -gt 0 -and ($cols -contains 'sheet_package_id')) {
            $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
            $clauses += "CAST(sheet_package_id AS NVARCHAR(36)) IN ($inList)"
        }
        if ($clauses.Count -eq 0) { continue }
        $rows = _QDM-RowsFromQuery -Sql "SELECT TOP (20) $select FROM [$table] WHERE $($clauses -join ' OR ')" -Parameters $params
        if ($rows.Count -gt 0) {
            [void]$sourceTables.Add($table)
            foreach ($row in $rows) {
                $recipients += (@{ source = $table } + $row)
            }
        }
    }

    $assessments = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($tr in $transitions) {
        if ([string]$tr.transition_type -ne 'STATE_CHANGE') { continue }
        $toState = $tr.to_value
        $sentFlag = $tr.notification_sent
        $matching = @($notifications | Where-Object {
            ($_.transition_id -eq $tr.id) -or ($toState -and ([string]$_.subject -like "*$toState*"))
        })
        $matchingJobs = @($jobs | Where-Object { $_.dedupe_key -and ([string]$tr.notification_id -like "*$($_.dedupe_key)*") })
        $outcome = 'not_queued'
        if ($matching | Where-Object { $_.success -in @($true, 1, '1', 'True') }) { $outcome = 'sent' }
        elseif ($matching | Where-Object { $_.success -in @($false, 0, '0', 'False') }) { $outcome = 'logged_but_failed' }
        elseif ($matchingJobs | Where-Object { [string]$_.status -in @('failed', 'dead') }) { $outcome = 'queued_but_failed' }
        elseif ($matchingJobs.Count -gt 0) { $outcome = 'queued' }
        elseif ($sentFlag -in @($true, 1, '1', 'True')) { $outcome = 'transition_marked_sent' }
        $assessments.Add(@{
            transition_id = $tr.id
            to_value = $toState
            notification_sent_flag = $sentFlag
            outcome = $outcome
        })
    }

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        notification_log = $notifications
        qc_notification_jobs = $jobs
        transition_rows = $transitions
        recipient_fields = $recipients
        transition_assessments = @($assessments)
    } -Warnings $warnings -SourceTables @($sourceTables) -QueryAssumptions @(
        'QC_NOTIFICATION jobs filtered by job_type when column exists.',
        'Outcome is heuristic based on available log/job/transition rows.'
    )
}

function Get-QCDebugDataIntegrityReport {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $membersResult = Get-QCDebugSheetPackageMembers @lookupParams
    $lookup = $membersResult.data.lookup
    $warnings = @()
    if ($membersResult.warnings) { $warnings += @($membersResult.warnings) }
    $members = $membersResult.data.members
    $issues = [System.Collections.Generic.List[hashtable]]::new()
    $sourceTables = @()
    if ($membersResult.source_tables) { $sourceTables += @($membersResult.source_tables) }

    $indexRows = @($members.sheet_index_rows)
    $pkgRows = @($members.sheet_packages_rows)
    $docRows = @($members.sheet_documents_rows)

    if ($pkgRows.Count -eq 0) {
        $issues.Add(@{ code = 'missing_package'; message = 'No sheet_packages row found for lookup.' })
    }
    if ($docRows.Count -eq 0) {
        $issues.Add(@{ code = 'missing_sheet_documents'; message = 'No sheet_documents rows found for lookup.' })
    }

    $expectedRoles = @('dgn', 'sheet_pdf', 'qc_pdf')
    $foundRoles = @($docRows | ForEach-Object { [string]$_.document_role } | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
    $lanePdfRows = @($members.sheet_package_qc_pdf_rows | Where-Object {
        $_.is_active -is [DBNull] -or $_.is_active -eq $true -or [string]$_.is_active -eq '1'
    })
    $hasLaneRegistry = ($lanePdfRows.Count -gt 0)
    if (-not $hasLaneRegistry -and $pkg) {
        foreach ($col in @('qc_pdf_guid', 'qc_chk_pdf_guid', 'qc_rev_pdf_guid')) {
            if ($pkg.$col -and -not ($pkg.$col -is [DBNull])) { $hasLaneRegistry = $true; break }
        }
    }
    foreach ($role in @($expectedRoles | Where-Object { $_ -notin $foundRoles })) {
        if ($role -eq 'qc_pdf' -and -not $hasLaneRegistry) { continue }
        $issues.Add(@{ code = 'missing_role'; message = "sheet_documents missing role $role."; role = $role })
    }

    $nameToGuids = @{}
    foreach ($row in $indexRows) {
        $name = [string]$row.document_name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $key = $name.ToLowerInvariant()
        if (-not $nameToGuids.ContainsKey($key)) { $nameToGuids[$key] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
        if ($row.document_guid) { [void]$nameToGuids[$key].Add([string]$row.document_guid) }
    }
    foreach ($entry in $nameToGuids.GetEnumerator()) {
        if ($entry.Value.Count -gt 1) {
            $issues.Add(@{
                code = 'duplicate_active_rows'
                message = "sheet_index has $($entry.Value.Count) GUIDs for document name $($entry.Key)."
                document_name = $entry.Key
                guids = @($entry.Value)
            })
            $warnings += _QDM-BuildWarning -Message "Duplicate sheet_index rows for $($entry.Key)" -Table 'sheet_index' -Column 'document_name'
        }
    }

    $pkg = if ($pkgRows.Count -gt 0) { $pkgRows[0] } else { @{} }
    $states = @{}
    if ($pkg.pw_state_name) { $states['sheet_packages'] = [string]$pkg.pw_state_name }
    foreach ($row in $docRows) {
        $role = [string]$row.document_role
        if ([string]::IsNullOrWhiteSpace($role)) { $role = 'unknown' }
        if ($row.pw_state_name) { $states["sheet_documents:$role"] = [string]$row.pw_state_name }
    }
    $uniqueStates = @($states.Values | Select-Object -Unique)
    if ($uniqueStates.Count -gt 1) {
        $issues.Add(@{ code = 'inconsistent_states'; message = 'Package/member states disagree.'; states = $states })
    }

    foreach ($row in $indexRows) {
        if (-not $row.sheet_package_id) {
            $issues.Add(@{
                code = 'missing_package_link'
                message = 'sheet_index row missing sheet_package_id.'
                document_guid = $row.document_guid
                document_name = $row.document_name
            })
        }
    }

    if (_QDM-TestTableExists -TableName 'v_sheet_package_status') {
        $sourceTables += 'v_sheet_package_status'
        $vcols = _QDM-GetColumns -TableName 'v_sheet_package_status'
        $viewRows = @()
        if ($lookup.lookup_type -eq 'package_id') {
            if ($vcols -contains 'sheet_package_id') {
                $viewRows = _QDM-RowsFromQuery -Sql "SELECT TOP (20) * FROM [v_sheet_package_status] WHERE CAST(sheet_package_id AS NVARCHAR(36)) = @pkg" -Parameters @{ pkg = $lookup.lookup_value }
            }
        }
        elseif ($lookup.lookup_type -eq 'document_path') {
            if (($vcols -contains 'folder_path') -and ($vcols -contains 'sheet_stem')) {
                $viewRows = _QDM-RowsFromQuery -Sql "SELECT TOP (20) * FROM [v_sheet_package_status] WHERE folder_path = @folderPath AND sheet_stem = @sheetStem" -Parameters @{
                    folderPath = $lookup.folder_path; sheetStem = $lookup.sheet_stem
                }
            }
        }
        else {
            $textCols = @($vcols | Where-Object { $_ -match '(?i)name|stem' })
            $where = _QDM-BuildLikeWhere -TableName 'v_sheet_package_status' -Columns $textCols -LikeParam (_QDM-LikePattern -Text $lookup.lookup_value)
            if ($where) {
                $viewRows = _QDM-RowsFromQuery -Sql "SELECT TOP (20) * FROM [v_sheet_package_status] WHERE $($where.clause)" -Parameters $where.params
            }
        }
        if ($viewRows.Count -eq 0 -and $pkgRows.Count -gt 0) {
            $issues.Add(@{ code = 'view_package_gap'; message = 'sheet_packages row exists but v_sheet_package_status returned no matches.' })
        }
    }

    $qcProcessType = if ($members.qc_process_type) { $members.qc_process_type } else {
        _QDM-BuildQcProcessTypeDiagnostics -Lookup $lookup -Members $members -ProjectWiseAvailable $false
    }
    foreach ($check in @($qcProcessType.checks | Where-Object { -not $_.passed })) {
        $issues.Add(@{
            code = [string]$check.code
            message = [string]$check.message
            document_guid = $check.document_guid
            document_name = $check.document_name
            expected = $check.expected
            actual = $check.actual
        })
    }

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        issues = @($issues)
        issue_count = $issues.Count
        qc_process_type = $qcProcessType
        members_snapshot = @{
            sheet_packages = $pkgRows
            sheet_documents = $docRows
            sheet_index = $indexRows
        }
    } -Warnings $warnings -SourceTables @($sourceTables | Select-Object -Unique) -QueryAssumptions @(
        'Integrity checks use only tables present in the database.'
    )
}

function Compare-QCProjectWiseToDatabase {
    <#
    .SYNOPSIS
    Read-only comparison of ProjectWise live document state vs QC_Pipeline telemetry.
    #>
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$SheetName = '',
        [string]$DocumentPath = ''
    )

    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    $membersResult = Get-QCDebugSheetPackageMembers @lookupParams
    $lookup = $membersResult.data.lookup
    $warnings = @()
    if ($membersResult.warnings) { $warnings += @($membersResult.warnings) }
    $members = $membersResult.data.members
    $comparisons = [System.Collections.Generic.List[hashtable]]::new()
    $missingInPw = [System.Collections.Generic.List[hashtable]]::new()
    $missingInDb = [System.Collections.Generic.List[hashtable]]::new()

    $dbDocs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($row in @($members.sheet_documents_rows)) {
        if ($row.document_guid) {
            $dbDocs.Add(@{
                role = $row.document_role
                document_guid = [string]$row.document_guid
                document_name = [string]$row.document_name
                pw_state_name = [string]$row.pw_state_name
                source = 'sheet_documents'
            })
        }
    }
    $pkgRowsLocal = @()
    if ($members.sheet_packages_rows) { $pkgRowsLocal = @($members.sheet_packages_rows) }
    $pkg = @{}
    if ($pkgRowsLocal.Count -gt 0 -and $null -ne $pkgRowsLocal[0]) { $pkg = $pkgRowsLocal[0] }
    foreach ($pair in @(
        @{ role = 'dgn'; guid = 'dgn_guid'; name = 'dgn_name'; state = 'pw_state_name' }
        @{ role = 'sheet_pdf'; guid = 'sheet_pdf_guid'; name = 'sheet_pdf_name'; state = 'pw_state_name' }
        @{ role = 'qc_pdf'; guid = 'qc_pdf_guid'; name = 'qc_pdf_name'; state = 'pw_state_name' }
    )) {
        if ($pkg[$pair.guid]) {
            $existing = @($dbDocs | Where-Object { $_.document_guid -ieq [string]$pkg[$pair.guid] })
            if ($existing.Count -eq 0) {
                $dbDocs.Add(@{
                    role = $pair.role
                    document_guid = [string]$pkg[$pair.guid]
                    document_name = [string]$pkg[$pair.name]
                    pw_state_name = [string]$pkg[$pair.state]
                    source = 'sheet_packages'
                })
            }
        }
    }

    if ($dbDocs.Count -eq 0) {
        $warnings += _QDM-BuildWarning -Message 'No document GUIDs found in database for comparison.'
        return _QDM-ToolResult -Data @{
            lookup = $lookup
            comparisons = @()
            missing_in_projectwise = @()
            missing_in_database = @()
            database_document_count = 0
            projectwise_available = $false
        } -Warnings $warnings -QueryAssumptions @('Database had no resolvable members; ProjectWise was not queried.')
    }

    $cfg = _QDM-Config
    $pw = $cfg.projectWise
    $ds = [string]$pw.datasourceName
    $credPath = [string]$pw.credentialPath
    $pwAvailable = $false
    $pwStateMap = @{}
    $pwNameMap = @{}

    if ([string]::IsNullOrWhiteSpace($ds) -or [string]::IsNullOrWhiteSpace($credPath)) {
        $warnings += _QDM-BuildWarning -Message 'projectWise.datasourceName or credentialPath missing; skipping live PW reads.'
    }
    elseif (-not (Get-Command -Name 'Invoke-PWAuthenticatedCommand' -ErrorAction SilentlyContinue)) {
        $warnings += _QDM-BuildWarning -Message 'Invoke-PWAuthenticatedCommand unavailable; run on a host with ProjectWise PowerShell.'
    }
    else {
        $guids = @($dbDocs | ForEach-Object { $_.document_guid } | Where-Object { Test-PWValidDocumentGuid -DocumentGuid $_ } | Select-Object -Unique)
        $modulesRoot = Split-Path -Parent $PSScriptRoot
        try {
            $pwResult = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -KeepSession -ScriptBlock {
                Import-Module (Join-Path $modulesRoot 'ProjectWise/PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue | Out-Null
                $states = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids
                $names = @{}
                $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
                if ($guidCmd) {
                    foreach ($doc in @(& $guidCmd -DocumentGUIDs $guids -ErrorAction SilentlyContinue)) {
                        $g = ''
                        try { $g = [string]$doc.DocumentGUID } catch { }
                        if ($g) { $names[$g.ToLowerInvariant()] = (Get-PWDocName -Doc $doc) }
                    }
                }
                return @{ states = $states; names = $names }
            }
            $pwStateMap = if ($pwResult.states) { $pwResult.states } else { @{} }
            $pwNameMap = if ($pwResult.names) { $pwResult.names } else { @{} }
            $pwAvailable = $true
        } catch {
            $warnings += _QDM-BuildWarning -Message "ProjectWise read failed: $($_.Exception.Message)"
            $pwStateMap = @{}
            $pwNameMap = @{}
        }
    }

    $warningList = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($warnings)) { $warningList.Add($w) }
    $folderPath = _QDM-ResolveMemberFolderPath -Lookup $lookup -PkgRow $pkg -IndexRows @($members.sheet_index_rows)
    $pwProcessTypeMap = _QDM-ReadProjectWiseProcessTypesForDocuments -Config $cfg -FolderPath $folderPath -Documents @($dbDocs) -Warnings $warningList
    $warnings = @($warningList)
    $qcProcessType = _QDM-BuildQcProcessTypeDiagnostics -Lookup $lookup -Members $members -PwProcessTypeByGuid $pwProcessTypeMap -ProjectWiseAvailable ($pwProcessTypeMap.Count -gt 0)
    $qcByGuid = @{}
    foreach ($row in @($qcProcessType.documents)) {
        if ($row.document_guid) { $qcByGuid[[string]$row.document_guid.ToLowerInvariant()] = $row }
    }

    foreach ($doc in $dbDocs) {
        $guid = [string]$doc.document_guid
        $guidKey = $guid.ToLowerInvariant()
        $pwState = if ($pwStateMap.ContainsKey($guidKey)) { [string]$pwStateMap[$guidKey] } else { $null }
        $pwName = if ($pwNameMap.ContainsKey($guidKey)) { [string]$pwNameMap[$guidKey] } else { $null }
        $dbState = [string]$doc.pw_state_name
        $qcRow = if ($qcByGuid.ContainsKey($guidKey)) { $qcByGuid[$guidKey] } else { @{} }
        $pwProcessRaw = if ($qcRow.projectwise_process_type) { [string]$qcRow.projectwise_process_type } else { $null }
        $pwProcessNorm = if ($qcRow.projectwise_process_type_normalized) { [string]$qcRow.projectwise_process_type_normalized } else { $null }
        $filenameInferred = if ($qcRow.filename_inferred_process_type) { [string]$qcRow.filename_inferred_process_type } else { $null }
        $dbProcessNorm = if ($qcRow.database_sheet_index_process_type) { [string]$qcRow.database_sheet_index_process_type } else { $null }
        $laneRegistryProcess = if ($qcRow.database_lane_registry_process_type) { [string]$qcRow.database_lane_registry_process_type } else { $null }
        $processMatch = $null
        if ($filenameInferred) {
            $processMatch = ($laneRegistryProcess -eq $filenameInferred) -and ((-not $pwProcessNorm) -or ($pwProcessNorm -eq $filenameInferred))
        } elseif ($dbProcessNorm -and $pwProcessNorm) {
            $processMatch = ($dbProcessNorm -eq $pwProcessNorm)
        }
        $mismatch = @{
            role = $doc.role
            document_guid = $guid
            database_name = $doc.document_name
            projectwise_name = $pwName
            database_state = $dbState
            projectwise_state = $pwState
            database_qc_process_type = $dbProcessNorm
            database_lane_registry_process_type = $laneRegistryProcess
            projectwise_qc_process_type = $pwProcessRaw
            projectwise_qc_process_type_normalized = $pwProcessNorm
            filename_inferred_process_type = $filenameInferred
            qc_process_type_match = $processMatch
            name_match = if ($pwName) { ($pwName -ieq $doc.document_name) } else { $null }
            state_match = if ($pwState) { ($pwState -ieq $dbState) } else { $null }
            found_in_projectwise = [bool]$pwState
        }
        if (-not $pwState -and $pwAvailable) {
            $missingInPw.Add($mismatch)
        }
        $comparisons.Add($mismatch)
        if ($pwState -and [string]::IsNullOrWhiteSpace($dbState)) {
            $missingInDb.Add($mismatch)
        }
    }

    return _QDM-ToolResult -Data @{
        lookup = $lookup
        comparisons = @($comparisons)
        missing_in_projectwise = @($missingInPw)
        missing_in_database = @($missingInDb)
        database_document_count = $dbDocs.Count
        projectwise_available = $pwAvailable
        qc_process_type = $qcProcessType
    } -Warnings $warnings -QueryAssumptions @(
        'Read-only Get-PWDocumentsByGUIDs workflow state lookup.',
        'Lane PDF qc_process_type is validated against *-prod/*-chk/*-rev filename suffixes.',
        'Stem/DGN qc_process_type compares sheet_index with live ProjectWise QC_Process_Type.',
        'No ProjectWise writes or arbitrary cmdlets are executed.'
    )
}

function _QDM-AutomationEventsAvailable {
    return (_QDM-TestTableExists -TableName 'automation_events')
}

function _QDM-GetJsonLogDirectories {
    $dirs = [System.Collections.Generic.List[string]]::new()
    if ($env:QC_JSON_LOG_DIR -and -not [string]::IsNullOrWhiteSpace($env:QC_JSON_LOG_DIR)) {
        $dirs.Add([string]$env:QC_JSON_LOG_DIR)
    }
    try {
        $cfg = _QDM-Config
        $settings = Get-QCAutomationTelemetrySettings -Config $cfg -ErrorAction SilentlyContinue
        if ($settings -and $settings.jsonLogDir -and -not [string]::IsNullOrWhiteSpace($settings.jsonLogDir)) {
            $dirs.Add([string]$settings.jsonLogDir)
        }
    } catch { }
    return @($dirs | Select-Object -Unique)
}

function _QDM-ParseJsonlEventLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
    try {
        $obj = $Line | ConvertFrom-Json
        $data = @{}
        if ($obj.data) {
            $h = ConvertTo-HashtableDeep -Value $obj.data -ErrorAction SilentlyContinue
            if ($h -is [hashtable]) { $data = $h }
        }
        return @{
            ts = [string]$obj.ts
            level = [string]$obj.level
            code = [string]$obj.code
            message = [string]$obj.message
            data = $data
            data_json = $Line
        }
    } catch { return $null }
}

function _QDM-ReadJsonlAutomationEvents {
    param(
        [string[]]$FilePatterns = @('Run-QCProcessor_*.jsonl', 'Watch-QCTrigger_*.jsonl', '*_*.jsonl'),
        [int]$MaxLines = 500,
        [hashtable]$Filter = @{},
        [switch]$ErrorsOnly
    )
    $events = [System.Collections.Generic.List[hashtable]]::new()
    $dirs = _QDM-GetJsonLogDirectories
    if ($dirs.Count -eq 0) { return @() }

    foreach ($dir in $dirs) {
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $files = @()
        foreach ($pat in $FilePatterns) {
            $files += @(Get-ChildItem -LiteralPath $dir -Filter $pat -File -ErrorAction SilentlyContinue)
        }
        $files = @($files | Sort-Object LastWriteTime -Descending | Select-Object -Unique FullName)
        foreach ($f in $files) {
            try {
                $lines = @(Get-Content -LiteralPath $f.FullName -Tail $MaxLines -ErrorAction SilentlyContinue)
                foreach ($line in $lines) {
                    $evt = _QDM-ParseJsonlEventLine -Line $line
                    if (-not $evt) { continue }
                    if ($ErrorsOnly -and $evt.level -notmatch '^(?i)(warning|error)$') { continue }
                    if ($Filter.ContainsKey('code') -and $Filter.code -and $evt.code -ne $Filter.code) { continue }
                    if ($Filter.ContainsKey('job_id') -and $Filter.job_id) {
                        $jid = if ($evt.data.jobId) { [string]$evt.data.jobId } elseif ($evt.data.job_id) { [string]$evt.data.job_id } else { '' }
                        if ($jid -ne [string]$Filter.job_id) { continue }
                    }
                    if ($Filter.ContainsKey('document_guid') -and $Filter.document_guid) {
                        $dg = if ($evt.data.documentGuid) { [string]$evt.data.documentGuid } elseif ($evt.data.document_guid) { [string]$evt.data.document_guid } else { '' }
                        if ($dg -ne [string]$Filter.document_guid) { continue }
                    }
                    if ($Filter.ContainsKey('sheet_package_id') -and $Filter.sheet_package_id) {
                        $pkg = if ($evt.data.sheetPackageId) { [string]$evt.data.sheetPackageId } elseif ($evt.data.sheet_package_id) { [string]$evt.data.sheet_package_id } else { '' }
                        if ($pkg -ne [string]$Filter.sheet_package_id) { continue }
                    }
                    $evt['source_file'] = $f.Name
                    $events.Add($evt)
                    if ($events.Count -ge $MaxLines) { return @($events) }
                }
            } catch { }
        }
    }
    return @($events)
}

function _QDM-QueryAutomationView {
    param(
        [Parameter(Mandatory)][string]$ViewName,
        [string]$WhereSql = '',
        [hashtable]$Parameters = @{},
        [int]$Limit = 200,
        [string]$OrderBy = 'ts DESC'
    )
    if (-not (_QDM-AutomationEventsAvailable)) {
        return @{ rows = @(); available = $false }
    }
    $lim = _QDM-SafeTopLimit -Limit $Limit
    $sql = "SELECT TOP ($lim) * FROM $ViewName"
    if ($WhereSql) { $sql += " WHERE $WhereSql" }
    if ($OrderBy) { $sql += " ORDER BY $OrderBy" }
    $rows = _QDM-RowsFromQuery -Sql $sql -Parameters $Parameters
    return @{ rows = @($rows); available = $true; view = $ViewName }
}

function Get-QCDebugRecentErrors {
    [CmdletBinding()]
    param(
        [int]$Limit = 100,
        [int]$Hours = 168,
        [switch]$ForceJsonlFallback
    )
    $warnings = @()
    $sourceTables = @()
    $fallback = $false
    $rows = @()

    if (-not $ForceJsonlFallback -and (_QDM-AutomationEventsAvailable)) {
        $sourceTables += 'v_mcp_recent_errors'
        $q = _QDM-QueryAutomationView -ViewName 'v_mcp_recent_errors' -Limit $Limit -WhereSql 'ts >= DATEADD(hour, -@hours, SYSDATETIMEOFFSET())' -Parameters @{ hours = $Hours }
        $rows = @($q.rows)
    } else {
        $fallback = $true
        $warnings += (_QDM-BuildWarning -Message 'Reading JSONL log files (automation_events unavailable or fallback requested).')
        $rows = @(_QDM-ReadJsonlAutomationEvents -MaxLines $Limit -ErrorsOnly)
    }

    return _QDM-ToolResult -Data @{
        fallback_mode = $fallback
        hours = $Hours
        event_count = $rows.Count
        events = $rows
    } -Warnings $warnings -SourceTables $sourceTables -QueryAssumptions @(
        'Primary source is automation_events via v_mcp_recent_errors.',
        'JSONL fallback scans QC_JSON_LOG_DIR / telemetry.automationEvents.jsonLogDir only when requested or DB unavailable.'
    )
}

function Get-QCDebugProcessHealth {
    [CmdletBinding()]
    param(
        [switch]$ForceJsonlFallback
    )
    $warnings = @()
    $sourceTables = @()
    $fallback = $false
    $rows = @()

    if (-not $ForceJsonlFallback -and (_QDM-AutomationEventsAvailable)) {
        $sourceTables += 'v_mcp_process_health'
        $q = _QDM-QueryAutomationView -ViewName 'v_mcp_process_health' -Limit 50 -OrderBy 'last_event_at DESC'
        $rows = @($q.rows)
        if ($rows.Count -eq 0) {
            $fallback = $true
            $warnings += (_QDM-BuildWarning -Message 'automation_events returned no rows; falling back to JSONL tail.')
            $events = @(_QDM-ReadJsonlAutomationEvents -MaxLines 2000)
            $groups = $events | Group-Object { if ($_.data.process_name) { $_.data.process_name } else { 'unknown' } }
            $rows = @($groups | ForEach-Object {
                @{
                    process_name = $_.Name
                    last_event_at = ($_.Group | Sort-Object { [string]$_.ts } -Descending | Select-Object -First 1).ts
                    error_count_24h = @($_.Group | Where-Object { $_.level -match '^(?i)error$' }).Count
                    warning_count_24h = @($_.Group | Where-Object { $_.level -match '^(?i)warning$' }).Count
                    event_count_24h = $_.Count
                }
            })
        }
    } else {
        $fallback = $true
        $warnings += (_QDM-BuildWarning -Message 'Process health derived from JSONL tail (limited accuracy).')
        $events = @(_QDM-ReadJsonlAutomationEvents -MaxLines 2000)
        $groups = $events | Group-Object { if ($_.data.process_name) { $_.data.process_name } else { 'unknown' } }
        $rows = @($groups | ForEach-Object {
            @{
                process_name = $_.Name
                last_event_at = ($_.Group | Sort-Object { [string]$_.ts } -Descending | Select-Object -First 1).ts
                error_count_24h = @($_.Group | Where-Object { $_.level -match '^(?i)error$' }).Count
                warning_count_24h = @($_.Group | Where-Object { $_.level -match '^(?i)warning$' }).Count
                event_count_24h = $_.Count
            }
        })
    }

    return _QDM-ToolResult -Data @{
        fallback_mode = $fallback
        processes = $rows
    } -Warnings $warnings -SourceTables $sourceTables -QueryAssumptions @(
        'v_mcp_process_health aggregates automation_events over the last 24 hours per process_name.'
    )
}

function Get-QCDebugAuditScanHistory {
    [CmdletBinding()]
    param(
        [int]$Limit = 200,
        [int]$Hours = 72,
        [switch]$ForceJsonlFallback
    )
    $warnings = @()
    $sourceTables = @()
    $fallback = $false
    $rows = @()

    if (-not $ForceJsonlFallback -and (_QDM-AutomationEventsAvailable)) {
        $sourceTables += 'v_mcp_audit_scan_history'
        $q = _QDM-QueryAutomationView -ViewName 'v_mcp_audit_scan_history' -Limit $Limit `
            -WhereSql 'ts >= DATEADD(hour, -@hours, SYSDATETIMEOFFSET())' -Parameters @{ hours = $Hours }
        $rows = @($q.rows)
    } else {
        $fallback = $true
        $warnings += (_QDM-BuildWarning -Message 'Audit scan history read from JSONL fallback.')
        $all = @(_QDM-ReadJsonlAutomationEvents -MaxLines ($Limit * 3))
        $rows = @($all | Where-Object {
            $_.code -like 'WATCH_AUDIT_*' -or $_.code -like 'AUDIT_EVENTS_*' -or $_.code -like 'AUDIT_*'
        } | Select-Object -First $Limit)
    }

    return _QDM-ToolResult -Data @{
        fallback_mode = $fallback
        hours = $Hours
        event_count = $rows.Count
        events = $rows
    } -Warnings $warnings -SourceTables $sourceTables -QueryAssumptions @(
        'Audit scan codes: WATCH_AUDIT_*, AUDIT_EVENTS_*, AUDIT_*.'
    )
}

function Get-QCDebugJobTimeline {
    [CmdletBinding()]
    param(
        [string]$JobId = '',
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$DocumentPath = '',
        [int]$Limit = 200,
        [switch]$ForceJsonlFallback
    )
    $warnings = @()
    $sourceTables = @('processing_jobs')
    $fallback = $false
    $resolvedJobId = $JobId

    if (-not $resolvedJobId) {
        $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
        if ($lookupParams.Keys.Count -gt 0) {
            $lookup = Resolve-QCDebugLookup @lookupParams
            if (_QDM-TestTableExists -TableName 'processing_jobs') {
                $filters = @()
                $params = @{}
                if ($lookup.sheet_package_ids.Count -gt 0) {
                    $inList = (@($lookup.sheet_package_ids) | ForEach-Object { "'$_'" }) -join ','
                    $filters += "sheet_package_id IN ($inList)"
                }
                if ($lookup.document_guids.Count -gt 0) {
                    $inList = (@($lookup.document_guids) | ForEach-Object { "'$($_.ToLowerInvariant())'" }) -join ','
                    $filters += "LOWER(source_path) LIKE '%' + LOWER(@docGuid) + '%'"
                    $params['docGuid'] = [string]$lookup.document_guids[0]
                }
                if ($filters.Count -gt 0) {
                    $jobRows = _QDM-RowsFromQuery -Sql "SELECT TOP (1) job_id FROM processing_jobs WHERE $($filters -join ' OR ') ORDER BY created_at DESC" -Parameters $params
                    if ($jobRows.Count -gt 0) { $resolvedJobId = [string]$jobRows[0].job_id }
                }
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($resolvedJobId)) {
        throw 'Provide job_id or a resolvable sheet/document/package lookup.'
    }

    $rows = @()
    $jobRow = $null
    if (_QDM-TestTableExists -TableName 'processing_jobs') {
        $jobCols = _QDM-SelectExistingColumns -TableName 'processing_jobs' -Requested @(
            'job_id', 'job_type', 'status', 'started_at', 'completed_at', 'duration_ms',
            'source_folder', 'source_path', 'error_code', 'error_message',
            'worker_machine_name', 'worker_pid', 'sheet_package_id'
        )
        if ($jobCols.Count -gt 0) {
            $select = ($jobCols | ForEach-Object { "[$_]" }) -join ', '
            $found = _QDM-RowsFromQuery -Sql "SELECT TOP (1) $select FROM [processing_jobs] WHERE job_id = @jobId" -Parameters @{ jobId = $resolvedJobId }
            if ($found.Count -gt 0) { $jobRow = $found[0] }
        }
    }
    if (-not $ForceJsonlFallback -and (_QDM-AutomationEventsAvailable)) {
        $sourceTables += 'v_mcp_job_timeline'
        $q = _QDM-QueryAutomationView -ViewName 'v_mcp_job_timeline' -Limit $Limit `
            -WhereSql 'job_id = @jobId' -Parameters @{ jobId = $resolvedJobId }
        $rows = @($q.rows)
    } else {
        $fallback = $true
        $warnings += (_QDM-BuildWarning -Message 'Job timeline read from JSONL fallback.')
        $rows = @(_QDM-ReadJsonlAutomationEvents -MaxLines ($Limit * 2) -Filter @{ job_id = $resolvedJobId })
    }

    return _QDM-ToolResult -Data @{
        fallback_mode = $fallback
        job_id = $resolvedJobId
        job = $jobRow
        event_count = $rows.Count
        events = $rows
    } -Warnings $warnings -SourceTables $sourceTables -QueryAssumptions @(
        'processing_jobs row (including worker_machine_name / worker_pid when present) plus automation events filtered by job_id.'
    )
}

function Get-QCDebugDocumentAutomationEvents {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$DocumentPath = '',
        [int]$Limit = 200,
        [switch]$ForceJsonlFallback
    )
    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    if ($lookupParams.Keys.Count -eq 0) {
        throw 'Provide one of: sheet_number, document_guid, package_id, document_path.'
    }
    $lookup = Resolve-QCDebugLookup @lookupParams
    $docGuid = $DocumentGuid
    if ([string]::IsNullOrWhiteSpace($docGuid) -and $lookup.document_guids.Count -gt 0) {
        $docGuid = [string]$lookup.document_guids[0]
    }
    if ([string]::IsNullOrWhiteSpace($docGuid)) {
        throw 'Could not resolve document_guid from lookup.'
    }

    $warnings = @()
    $sourceTables = @('v_mcp_document_debug_events')
    $fallback = $false
    $rows = @()

    if (-not $ForceJsonlFallback -and (_QDM-AutomationEventsAvailable)) {
        $q = _QDM-QueryAutomationView -ViewName 'v_mcp_document_debug_events' -Limit $Limit `
            -WhereSql 'document_guid = @documentGuid' -Parameters @{ documentGuid = $docGuid }
        $rows = @($q.rows)
    } else {
        $fallback = $true
        $warnings += (_QDM-BuildWarning -Message 'Document events read from JSONL fallback.')
        $rows = @(_QDM-ReadJsonlAutomationEvents -MaxLines ($Limit * 2) -Filter @{ document_guid = $docGuid })
    }

    return _QDM-ToolResult -Data @{
        fallback_mode = $fallback
        document_guid = $docGuid
        lookup = $lookup
        event_count = $rows.Count
        events = $rows
    } -Warnings $warnings -SourceTables $sourceTables -QueryAssumptions @(
        'Automation events with indexed document_guid column.'
    )
}

function Get-QCDebugPackageAutomationEvents {
    [CmdletBinding()]
    param(
        [string]$SheetNumber = '',
        [string]$DocumentGuid = '',
        [string]$PackageId = '',
        [string]$DocumentPath = '',
        [int]$Limit = 200,
        [switch]$ForceJsonlFallback
    )
    $lookupParams = _QDM-GetLookupBoundParameters -Bound $PSBoundParameters
    if ($lookupParams.Keys.Count -eq 0) {
        throw 'Provide one of: sheet_number, document_guid, package_id, document_path.'
    }
    $lookup = Resolve-QCDebugLookup @lookupParams
    $pkgId = $PackageId
    if ([string]::IsNullOrWhiteSpace($pkgId) -and $lookup.sheet_package_ids.Count -gt 0) {
        $pkgId = [string]$lookup.sheet_package_ids[0]
    }
    if ([string]::IsNullOrWhiteSpace($pkgId)) {
        throw 'Could not resolve sheet_package_id from lookup.'
    }

    $warnings = @()
    $sourceTables = @('v_mcp_package_debug_events')
    $fallback = $false
    $rows = @()

    if (-not $ForceJsonlFallback -and (_QDM-AutomationEventsAvailable)) {
        $q = _QDM-QueryAutomationView -ViewName 'v_mcp_package_debug_events' -Limit $Limit `
            -WhereSql 'sheet_package_id = @packageId' -Parameters @{ packageId = $pkgId }
        $rows = @($q.rows)
    } else {
        $fallback = $true
        $warnings += (_QDM-BuildWarning -Message 'Package events read from JSONL fallback.')
        $rows = @(_QDM-ReadJsonlAutomationEvents -MaxLines ($Limit * 2) -Filter @{ sheet_package_id = $pkgId })
    }

    return _QDM-ToolResult -Data @{
        fallback_mode = $fallback
        sheet_package_id = $pkgId
        lookup = $lookup
        event_count = $rows.Count
        events = $rows
    } -Warnings $warnings -SourceTables $sourceTables -QueryAssumptions @(
        'Automation events with indexed sheet_package_id column.'
    )
}

Export-ModuleMember -Function @(
    'Initialize-QCDebugMcpContext'
    'Resolve-QCDebugLookup'
    'Search-QCDebugSheet'
    'Get-QCDebugSheetIdentity'
    'Get-QCDebugSheetPackageMembers'
    'Get-QCDebugSheetTimeline'
    'Get-QCDebugNotificationDiagnostics'
    'Get-QCDebugDataIntegrityReport'
    'Get-QCDebugQcProcessTypeDiagnostics'
    'Compare-QCProjectWiseToDatabase'
    'Get-QCDebugRecentErrors'
    'Get-QCDebugProcessHealth'
    'Get-QCDebugAuditScanHistory'
    'Get-QCDebugJobTimeline'
    'Get-QCDebugDocumentAutomationEvents'
    'Get-QCDebugPackageAutomationEvents'
)
