# pw-qc-debug MCP server (PowerShell, read-only diagnostics).
# Communicates via MCP stdio transport (Content-Length framed JSON-RPC 2.0).

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesRoot = Join-Path $repoRoot 'modules'
Import-Module (Join-Path $modulesRoot 'QC.DebugMcp.psm1') -Force

$appSettings = $env:PWQC_APPSETTINGS
if ([string]::IsNullOrWhiteSpace($appSettings)) {
    $appSettings = Join-Path $repoRoot 'appsettings.json'
}
Initialize-QCDebugMcpContext -AppSettingsPath $appSettings | Out-Null

function Write-McpLog {
    param([Parameter(Mandatory)][string]$Message)
    [Console]::Error.WriteLine($Message)
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
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    [Console]::Out.Write("Content-Length: $($bytes.Length)`r`n`r`n")
    [Console]::Out.Write($json)
    [Console]::Out.Flush()
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
    (New-McpToolSchema -Name 'compare_projectwise_to_database' -Description 'Read-only comparison of live ProjectWise workflow state vs QC_Pipeline telemetry.')
)

$script:ToolDispatch = @{
    search_sheet = { param($a) Search-QCDebugSheet @a }
    get_sheet_identity = { param($a) Get-QCDebugSheetIdentity @a }
    get_sheet_package_members = { param($a) Get-QCDebugSheetPackageMembers @a }
    get_sheet_debug_timeline = { param($a) Get-QCDebugSheetTimeline @a }
    get_notification_diagnostics = { param($a) Get-QCDebugNotificationDiagnostics @a }
    get_data_integrity_report = { param($a) Get-QCDebugDataIntegrityReport @a }
    compare_projectwise_to_database = { param($a) Compare-QCProjectWiseToDatabase @a }
}

function Get-LookupArguments {
    param([hashtable]$Arguments)
    $args = @{}
    foreach ($key in @('sheet_number', 'document_guid', 'package_id', 'document_path', 'sheet_name', 'limit')) {
        if ($Arguments.ContainsKey($key) -and $null -ne $Arguments[$key] -and -not [string]::IsNullOrWhiteSpace([string]$Arguments[$key])) {
            $args[$key] = $Arguments[$key]
        }
    }
    if ($args.Keys.Count -eq 0 -or (-not ($args.ContainsKey('sheet_number') -or $args.ContainsKey('document_guid') -or $args.ContainsKey('package_id') -or $args.ContainsKey('document_path') -or $args.ContainsKey('sheet_name')))) {
        throw 'Provide one of: sheet_number, document_guid, package_id, document_path, or sheet_name.'
    }
    if ($args.ContainsKey('limit')) { $args['limit'] = [int]$args['limit'] }
    return $args
}

function Invoke-McpTool {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Arguments = @{}
    )
    if (-not $script:ToolDispatch.ContainsKey($Name)) {
        throw "Unknown tool: $Name"
    }
    $bound = Get-LookupArguments -Arguments $Arguments
    return & $script:ToolDispatch[$Name] $bound
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
    $params = @{}
    if ($Request.params) { $params = [hashtable]$Request.params }

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
                if ($params.arguments) { $toolArgs = [hashtable]$params.arguments }
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
        Write-McpLog "Request error ($method): $($_.Exception.Message)"
    }
}

Write-McpLog "pw-qc-debug MCP server started (appsettings: $appSettings)"

while ($true) {
    $raw = Read-McpMessage
    if ($null -eq $raw) { break }
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }
    try {
        $request = $raw | ConvertFrom-Json -AsHashtable
    } catch {
        Write-McpLog "Invalid JSON: $($_.Exception.Message)"
        continue
    }
    if (-not $request.method) { continue }
    Handle-McpRequest -Request $request
}
