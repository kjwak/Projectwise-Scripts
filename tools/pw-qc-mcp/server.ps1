# pw-qc-debug MCP server (PowerShell, read-only diagnostics).
# Communicates via MCP stdio transport (Content-Length framed JSON-RPC 2.0).

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesRoot = Join-Path $repoRoot 'modules'
$script:McpContextReady = $false

function Initialize-McpRuntime {
    if ($script:McpContextReady) { return }
    Import-Module (Join-Path $modulesRoot 'QC.DebugMcp.psm1') -Force -WarningAction SilentlyContinue | Out-Null
    Import-Module (Join-Path $modulesRoot 'Core.Runtime.psm1') -Force -WarningAction SilentlyContinue | Out-Null
    Import-Module (Join-Path $modulesRoot 'Core.Telemetry.psm1') -Force -WarningAction SilentlyContinue | Out-Null
    $appSettings = $env:PWQC_APPSETTINGS
    if ([string]::IsNullOrWhiteSpace($appSettings)) {
        $appSettings = Join-Path $repoRoot 'appsettings.json'
    }
    Initialize-QCDebugMcpContext -AppSettingsPath $appSettings | Out-Null
    $script:McpContextReady = $true
}

function Write-McpLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [switch]$IsError
    )
    if (-not $IsError) { return }
    [Console]::Error.WriteLine($Message)
}

function ConvertTo-McpHashtable {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($key in $Value.Keys) { $h[$key] = ConvertTo-McpHashtable $Value[$key] }
        return $h
    }
    if ($Value -is [string] -or $Value -is [System.ValueType] -or $Value -is [decimal]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $arr = @()
        foreach ($item in $Value) { $arr += ,(ConvertTo-McpHashtable $item) }
        return $arr
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-McpHashtable $p.Value }
        return $h
    }
    return $Value
}

function Read-McpMessage {
    $headers = @{}
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { return $null }
        if ($line -eq '') { break }
        $idx = $line.IndexOf(':')
        if ($idx -gt 0) {
            $name = $line.Substring(0, $idx).Trim().ToLowerInvariant()
            $value = $line.Substring($idx + 1).Trim()
            $headers[$name] = $value
        }
    }
    if (-not $headers.ContainsKey('content-length')) { return $null }
    $length = [int]$headers['content-length']
    $buffer = New-Object char[] $length
    $read = 0
    while ($read -lt $length) {
        $n = [Console]::In.Read($buffer, $read, $length - $read)
        if ($n -le 0) { break }
        $read += $n
    }
    return (-join $buffer[0..($read - 1)])
}

function Write-McpMessage {
    param([Parameter(Mandatory)]$Payload)
    $json = $Payload | ConvertTo-Json -Depth 80 -Compress
    $body = [System.Text.Encoding]::UTF8.GetBytes($json)
    $header = [System.Text.Encoding]::ASCII.GetBytes("Content-Length: $($body.Length)`r`n`r`n")
    $stdout = [Console]::OpenStandardOutput()
    $stdout.Write($header, 0, $header.Length)
    $stdout.Write($body, 0, $body.Length)
    $stdout.Flush()
}

function New-McpToolSchema {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Description,
        [hashtable]$ExtraProperties = @{},
        [string[]]$Required = @()
    )
    $lookupProps = @{
        sheet_number = @{ type = 'string'; description = 'Sheet number or document name fragment.' }
        document_guid = @{ type = 'string'; description = 'ProjectWise document GUID.' }
        package_id = @{ type = 'string'; description = 'sheet_package_id GUID.' }
        document_path = @{ type = 'string'; description = 'Full ProjectWise document path (pw:\\datasource\\Documents\\...\\file.pdf).' }
        sheet_name = @{ type = 'string'; description = 'Alias for sheet_number (backward compatible).' }
    }
    foreach ($k in $ExtraProperties.Keys) { $lookupProps[$k] = $ExtraProperties[$k] }
    return @{
        name = $Name
        description = $Description
        inputSchema = @{
            type = 'object'
            properties = $lookupProps
            required = @($Required)
        }
    }
}

$script:McpTools = @(
    (New-McpToolSchema -Name 'search_sheet' -Description 'Search likely sheet/package/document tables for a sheet number, document GUID, or package ID.'),
    (New-McpToolSchema -Name 'get_sheet_identity' -Description 'Resolve candidate sheet package IDs, document GUIDs, names, and roles.'),
    (New-McpToolSchema -Name 'get_sheet_package_members' -Description 'Return package members (DGN, sheet PDF, QC PDF) and cross-table consistency hints.'),
    (New-McpToolSchema -Name 'get_sheet_debug_timeline' -Description 'Build a combined timeline from available QC telemetry tables.' -ExtraProperties @{ limit = @{ type = 'integer'; description = 'Max events (default 200).'; default = 200 } }),
    (New-McpToolSchema -Name 'get_notification_diagnostics' -Description 'Diagnose notification queue/log outcomes for a sheet.' -ExtraProperties @{ limit = @{ type = 'integer'; description = 'Max rows per source (default 100).'; default = 100 } }),
    (New-McpToolSchema -Name 'get_data_integrity_report' -Description 'Compare package/document identity and flag stale or inconsistent rows.'),
    (New-McpToolSchema -Name 'compare_projectwise_to_database' -Description 'Read-only comparison of live ProjectWise workflow state vs QC_Pipeline telemetry.'),
    (New-McpToolSchema -Name 'get_recent_errors' -Description 'Recent warning/error automation events from automation_events (DB-first).' -ExtraProperties @{
        limit = @{ type = 'integer'; description = 'Max events (default 100).'; default = 100 }
        hours = @{ type = 'integer'; description = 'Lookback hours (default 168).'; default = 168 }
        force_jsonl_fallback = @{ type = 'boolean'; description = 'Force JSONL log fallback instead of database.'; default = $false }
    } -Required @()),
    (New-McpToolSchema -Name 'get_process_health' -Description 'Per-process automation health summary from automation_events.' -ExtraProperties @{
        force_jsonl_fallback = @{ type = 'boolean'; description = 'Force JSONL log fallback.'; default = $false }
    } -Required @()),
    (New-McpToolSchema -Name 'get_audit_scan_history' -Description 'Watcher audit scan events from automation_events.' -ExtraProperties @{
        limit = @{ type = 'integer'; description = 'Max events (default 200).'; default = 200 }
        hours = @{ type = 'integer'; description = 'Lookback hours (default 72).'; default = 72 }
        force_jsonl_fallback = @{ type = 'boolean'; description = 'Force JSONL log fallback.'; default = $false }
    } -Required @()),
    (New-McpToolSchema -Name 'get_job_timeline' -Description 'Automation event timeline for a processing job.' -ExtraProperties @{
        job_id = @{ type = 'string'; description = 'processing_jobs.job_id (optional if lookup provided).' }
        limit = @{ type = 'integer'; description = 'Max events (default 200).'; default = 200 }
        force_jsonl_fallback = @{ type = 'boolean'; description = 'Force JSONL log fallback.'; default = $false }
    }),
    (New-McpToolSchema -Name 'get_document_debug_events' -Description 'Automation events for a document GUID.' -ExtraProperties @{
        limit = @{ type = 'integer'; description = 'Max events (default 200).'; default = 200 }
        force_jsonl_fallback = @{ type = 'boolean'; description = 'Force JSONL log fallback.'; default = $false }
    }),
    (New-McpToolSchema -Name 'get_package_debug_events' -Description 'Automation events for a sheet package.' -ExtraProperties @{
        limit = @{ type = 'integer'; description = 'Max events (default 200).'; default = 200 }
        force_jsonl_fallback = @{ type = 'boolean'; description = 'Force JSONL log fallback.'; default = $false }
    })
)

$script:ToolDispatch = @{
    search_sheet = { param($a) Search-QCDebugSheet @a }
    get_sheet_identity = { param($a) Get-QCDebugSheetIdentity @a }
    get_sheet_package_members = { param($a) Get-QCDebugSheetPackageMembers @a }
    get_sheet_debug_timeline = { param($a) Get-QCDebugSheetTimeline @a }
    get_notification_diagnostics = { param($a) Get-QCDebugNotificationDiagnostics @a }
    get_data_integrity_report = { param($a) Get-QCDebugDataIntegrityReport @a }
    compare_projectwise_to_database = { param($a) Compare-QCProjectWiseToDatabase @a }
    get_recent_errors = { param($a) Get-QCDebugRecentErrors @a }
    get_process_health = { param($a) Get-QCDebugProcessHealth @a }
    get_audit_scan_history = { param($a) Get-QCDebugAuditScanHistory @a }
    get_job_timeline = { param($a) Get-QCDebugJobTimeline @a }
    get_document_debug_events = { param($a) Get-QCDebugDocumentAutomationEvents @a }
    get_package_debug_events = { param($a) Get-QCDebugPackageAutomationEvents @a }
}

function Convert-McpToolArguments {
    param([hashtable]$Arguments)
    $map = @{
        sheet_number = 'SheetNumber'
        document_guid = 'DocumentGuid'
        package_id = 'PackageId'
        document_path = 'DocumentPath'
        sheet_name = 'SheetName'
        job_id = 'JobId'
        limit = 'Limit'
        hours = 'Hours'
        force_jsonl_fallback = 'ForceJsonlFallback'
    }
    $out = @{}
    foreach ($k in $Arguments.Keys) {
        $target = if ($map.ContainsKey($k)) { $map[$k] } else { $k }
        $out[$target] = $Arguments[$k]
    }
    return $out
}

function Get-LookupArguments {
    param(
        [hashtable]$Arguments,
        [switch]$RequireLookup
    )
    $args = @{}
    foreach ($key in @('sheet_number', 'document_guid', 'package_id', 'document_path', 'sheet_name', 'limit', 'hours', 'job_id', 'force_jsonl_fallback')) {
        if ($Arguments.ContainsKey($key) -and $null -ne $Arguments[$key] -and -not [string]::IsNullOrWhiteSpace([string]$Arguments[$key])) {
            $args[$key] = $Arguments[$key]
        }
    }
    if ($Arguments.ContainsKey('force_jsonl_fallback')) {
        $args['force_jsonl_fallback'] = [bool]$Arguments['force_jsonl_fallback']
    }
    if ($RequireLookup -and ($args.Keys.Count -eq 0 -or -not ($args.ContainsKey('sheet_number') -or $args.ContainsKey('document_guid') -or $args.ContainsKey('package_id') -or $args.ContainsKey('document_path') -or $args.ContainsKey('sheet_name') -or $args.ContainsKey('job_id')))) {
        throw 'Provide one of: sheet_number, document_guid, package_id, document_path, sheet_name, or job_id.'
    }
    if ($args.ContainsKey('limit')) { $args['limit'] = [int]$args['limit'] }
    if ($args.ContainsKey('hours')) { $args['hours'] = [int]$args['hours'] }
    return $args
}

function Invoke-McpTool {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Arguments = @{}
    )
    Initialize-McpRuntime
    if (-not $script:ToolDispatch.ContainsKey($Name)) {
        throw "Unknown tool: $Name"
    }
    $noLookupTools = @('get_recent_errors', 'get_process_health', 'get_audit_scan_history')
    if ($noLookupTools -contains $Name) {
        $bound = Get-LookupArguments -Arguments $Arguments
    } elseif ($Name -eq 'get_job_timeline') {
        $bound = Get-LookupArguments -Arguments $Arguments
        if (-not $bound.ContainsKey('job_id') -and -not ($bound.ContainsKey('sheet_number') -or $bound.ContainsKey('document_guid') -or $bound.ContainsKey('package_id') -or $bound.ContainsKey('document_path') -or $bound.ContainsKey('sheet_name'))) {
            throw 'get_job_timeline requires job_id or a sheet/document/package lookup.'
        }
    } else {
        $bound = Get-LookupArguments -Arguments $Arguments -RequireLookup
    }
    return & $script:ToolDispatch[$Name] (Convert-McpToolArguments -Arguments $bound)
}

function Send-McpResponse {
    param(
        $Id,
        $Result = $null,
        $Error = $null
    )
    $msg = @{ jsonrpc = '2.0'; id = $Id }
    if ($null -ne $Error) { $msg.error = $Error } else { $msg.result = $Result }
    Write-McpMessage -Payload $msg
}

function Handle-McpRequest {
    param([Parameter(Mandatory)][hashtable]$Request)

    $method = [string]$Request.method
    $id = $Request.id
    if ($null -eq $id -and $method -like 'notifications/*') { return }

    $params = @{}
    if ($Request.ContainsKey('params') -and $Request.params) {
        $params = ConvertTo-McpHashtable $Request.params
    }

    try {
        switch ($method) {
            'initialize' {
                Send-McpResponse -Id $id -Result @{
                    protocolVersion = '2024-11-05'
                    capabilities = @{ tools = @{} }
                    serverInfo = @{ name = 'pw-qc-debug'; version = '1.0.0' }
                }
            }
            'tools/list' {
                Send-McpResponse -Id $id -Result @{ tools = $script:McpTools }
            }
            'tools/call' {
                $toolName = [string]$params.name
                $toolArgs = @{}
                if ($params.ContainsKey('arguments') -and $params.arguments) {
                    $toolArgs = ConvertTo-McpHashtable $params.arguments
                }
                $data = Invoke-McpTool -Name $toolName -Arguments $toolArgs
                $json = $data | ConvertTo-Json -Depth 80 -Compress
                Send-McpResponse -Id $id -Result @{
                    content = @(@{ type = 'text'; text = $json })
                    structuredContent = $data
                    isError = $false
                }
            }
            'ping' {
                Send-McpResponse -Id $id -Result @{}
            }
            default {
                if ($null -ne $id) {
                    Send-McpResponse -Id $id -Error @{
                        code = -32601
                        message = "Method not found: $method"
                    }
                }
            }
        }
    } catch {
        if ($method -eq 'tools/call' -and $null -ne $id) {
            Send-McpResponse -Id $id -Result @{
                content = @(@{ type = 'text'; text = $_.Exception.Message })
                isError = $true
            }
        }
        elseif ($null -ne $id) {
            Send-McpResponse -Id $id -Error @{
                code = -32000
                message = $_.Exception.Message
            }
        }
        Write-McpLog -IsError "Request error ($method): $($_.Exception.Message)"
    }
}

while ($true) {
    $raw = Read-McpMessage
    if ($null -eq $raw) { break }
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }
    try {
        $request = ConvertTo-McpHashtable ($raw | ConvertFrom-Json)
    } catch {
        Write-McpLog -IsError "Invalid JSON: $($_.Exception.Message)"
        continue
    }
    if (-not $request.ContainsKey('method') -or -not $request.method) { continue }
    Handle-McpRequest -Request $request
}
