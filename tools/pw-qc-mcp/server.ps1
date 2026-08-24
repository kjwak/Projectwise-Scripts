# pw-qc-debug MCP server (PowerShell, read-only diagnostics).
# Communicates via MCP stdio transport (newline-delimited JSON-RPC 2.0).
#
# SQL tools run on this stdio thread. ProjectWise tools are dispatched to
# pw_qc_worker.ps1 so Connect-PW cannot stall get_recent_errors / get_process_health.

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding $false
$script:McpUtf8 = New-Object System.Text.UTF8Encoding $false
$script:McpStdin = [Console]::OpenStandardInput()
$script:McpStdout = [Console]::OpenStandardOutput()
$script:McpWriteLock = New-Object object
$script:McpSync = [hashtable]::Synchronized(@{ PwBusy = $false })
$script:PwToolNames = @(
    'compare_projectwise_to_database'
    'warm_projectwise_session'
    'get_qc_process_type_diagnostics'
)

$repoRoot = $env:PWQC_REPO_ROOT
if ([string]::IsNullOrWhiteSpace($repoRoot)) {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$modulesRoot = Join-Path $repoRoot 'modules'
$scriptsRoot = Join-Path $repoRoot 'scripts'
$script:McpContextReady = $false

function Initialize-McpRuntime {
    if ($script:McpContextReady) { return }
    . (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
    Import-QCModuleBootstrapSet -FeatureModules @(
        'Diagnostics\QC.DebugMcp.psm1'
        'Core\Core.Telemetry.psm1'
    ) -RequiredCommands @(
        'Initialize-QCDebugMcpContext'
        'Get-QCAppSettingsConfig'
        'Search-QCDebugSheet'
    ) -Context 'pw-qc-mcp server'
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

function Read-McpLine {
    $bytes = New-Object System.Collections.Generic.List[byte]
    while ($true) {
        $b = $script:McpStdin.ReadByte()
        if ($b -lt 0) {
            if ($bytes.Count -eq 0) { return $null }
            break
        }
        if ($b -eq 10) { break }
        if ($b -ne 13) { $bytes.Add([byte]$b) }
    }
    return $script:McpUtf8.GetString($bytes.ToArray())
}

function Read-McpMessage {
    # MCP stdio is newline-delimited JSON. Content-Length (LSP) is accepted as a fallback.
    while ($true) {
        $line = Read-McpLine
        if ($null -eq $line) { return $null }
        if ($line -eq '') { continue }
        if ($line -match '^Content-Length:\s*(\d+)\s*$') {
            $length = [int]$Matches[1]
            while ($true) {
                $headerLine = Read-McpLine
                if ($null -eq $headerLine -or $headerLine -eq '') { break }
            }
            $buffer = New-Object byte[] $length
            $offset = 0
            while ($offset -lt $length) {
                $read = $script:McpStdin.Read($buffer, $offset, $length - $offset)
                if ($read -le 0) { break }
                $offset += $read
            }
            return $script:McpUtf8.GetString($buffer, 0, $offset)
        }
        return $line
    }
}

function Write-McpMessage {
    param([Parameter(Mandatory)]$Payload)
    if ($Payload -is [string]) {
        $json = $Payload.Trim()
    } else {
        $json = $Payload | ConvertTo-Json -Depth 80 -Compress
    }
    if ($json.Contains("`n") -or $json.Contains("`r")) {
        $json = ($json -replace '[\r\n]+', ' ')
    }
    $body = $script:McpUtf8.GetBytes($json)
    [System.Threading.Monitor]::Enter($script:McpWriteLock)
    try {
        $script:McpStdout.Write($body, 0, $body.Length)
        $script:McpStdout.WriteByte(10)
        $script:McpStdout.Flush()
    } finally {
        [System.Threading.Monitor]::Exit($script:McpWriteLock)
    }
}

function Send-McpJsonResult {
    param($Id, [Parameter(Mandatory)][string]$ResultJson)
    $idJson = if ($null -eq $Id) { 'null' } elseif ($Id -is [string]) { '"' + ($Id -replace '"','\"') + '"' } else { [string]$Id }
    Write-McpMessage -Payload "{`"jsonrpc`":`"2.0`",`"id`":$idJson,`"result`":$ResultJson}"
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
    (New-McpToolSchema -Name 'get_qc_process_type_diagnostics' -Description 'Compare qc_process_type across lane filenames (*-prod/*-chk/*-rev), sheet_index, lane registry, and ProjectWise.'),
    (New-McpToolSchema -Name 'compare_projectwise_to_database' -Description 'Read-only comparison of live ProjectWise workflow state vs QC_Pipeline telemetry.'),
    (New-McpToolSchema -Name 'warm_projectwise_session' -Description 'Pre-connect ProjectWise in the MCP worker. Call before compare_projectwise_to_database to reduce timeout risk.' -Required @()),
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
    get_qc_process_type_diagnostics = { param($a) Get-QCDebugQcProcessTypeDiagnostics @a }
    compare_projectwise_to_database = { param($a) Compare-QCProjectWiseToDatabase @a }
    get_recent_errors = { param($a) Get-QCDebugRecentErrors @a }
    get_process_health = { param($a) Get-QCDebugProcessHealth @a }
    get_audit_scan_history = { param($a) Get-QCDebugAuditScanHistory @a }
    get_job_timeline = { param($a) Get-QCDebugJobTimeline @a }
    get_document_debug_events = { param($a) Get-QCDebugDocumentAutomationEvents @a }
    get_package_debug_events = { param($a) Get-QCDebugPackageAutomationEvents @a }
}

$script:ToolsListJson = ($script:McpTools | ConvertTo-Json -Depth 20 -Compress)
$script:EmptyResourcesJson = '{"resources":[]}'
$script:EmptyPromptsJson = '{"prompts":[]}'

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

function Stop-McpPwWorker {
    $proc = $script:McpSync.PwWorker
    $script:McpSync.PwWorker = $null
    $script:McpSync.PwBusy = $false
    if ($null -eq $proc) { return }
    try { if ($proc.StandardInput) { $proc.StandardInput.Close() } } catch { }
    try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
    try { [void]$proc.WaitForExit(3000) } catch { }
}

function Start-McpPwToolAsync {
    param(
        $McpId,
        [string]$ToolName,
        [hashtable]$ToolArgs
    )
    if ([bool]$script:McpSync.PwBusy) {
        Send-McpResponse -Id $McpId -Result @{
            content = @(@{ type = 'text'; text = 'ProjectWise MCP worker is busy. Retry this tool after the in-flight PW call finishes.' })
            isError = $true
        }
        return
    }
    $script:McpSync.PwBusy = $true
    $requestId = [guid]::NewGuid().ToString('N')
    $req = @{ id = $requestId; tool = $ToolName; arguments = $ToolArgs } | ConvertTo-Json -Depth 20 -Compress
    $appSettings = $env:PWQC_APPSETTINGS
    if ([string]::IsNullOrWhiteSpace($appSettings)) { $appSettings = Join-Path $repoRoot 'appsettings.json' }
    $envMap = [hashtable]::Synchronized(@{
        PWQC_REPO_ROOT = $repoRoot
        PWQC_APPSETTINGS = $appSettings
    })
    foreach ($key in @('PWQC_SQL_SERVER', 'PWQC_SQL_DATABASE', 'PWQC_SQL_TRUST_CERT', 'PWQC_SQL_DRIVER')) {
        $val = [Environment]::GetEnvironmentVariable($key)
        if ($val) { $envMap[$key] = $val }
    }

    $ps = [PowerShell]::Create()
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($Sync, $ReqJson, $RequestId, $Stdout, $Lock, $Utf8, $McpId, $TimeoutMs, $WorkerPs1, $RepoRoot, $EnvMap)
        function Write-McpPayload([string]$PayloadJson) {
            $body = $Utf8.GetBytes($PayloadJson)
            [System.Threading.Monitor]::Enter($Lock)
            try {
                $Stdout.Write($body, 0, $body.Length)
                $Stdout.WriteByte(10)
                $Stdout.Flush()
            } finally {
                [System.Threading.Monitor]::Exit($Lock)
            }
        }
        function New-McpToolErrorJson([string]$Text) {
            $idJson = $McpId | ConvertTo-Json -Compress
            return '{"jsonrpc":"2.0","id":' + $idJson + ',"result":{"content":[{"type":"text","text":' + ($Text | ConvertTo-Json -Compress) + '}],"isError":true}}'
        }
        try {
            $proc = $Sync.PwWorker
            if ($null -eq $proc -or $proc.HasExited) {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
                $psi.Arguments = "-NoLogo -MTA -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$WorkerPs1`" -Worker"
                $psi.WorkingDirectory = $RepoRoot
                $psi.RedirectStandardInput = $true
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                foreach ($key in @($EnvMap.Keys)) {
                    if ($EnvMap[$key]) { $psi.EnvironmentVariables[$key] = [string]$EnvMap[$key] }
                }
                $proc = [System.Diagnostics.Process]::Start($psi)
                $Sync.PwWorkerErrTask = $proc.StandardError.ReadToEndAsync()
                $readyTask = $proc.StandardOutput.ReadLineAsync()
                if (-not $readyTask.Wait(30000)) {
                    try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
                    $Sync.PwWorker = $null
                    Write-McpPayload (New-McpToolErrorJson 'ProjectWise MCP worker did not become ready within 30s.')
                    return
                }
                $readyDeadline = [datetime]::UtcNow.AddSeconds(30)
                $readyObj = $null
                $ready = [string]$readyTask.Result
                while ($true) {
                    if (-not [string]::IsNullOrWhiteSpace($ready) -and $ready.Trim().StartsWith('{')) {
                        try { $readyObj = $ready | ConvertFrom-Json } catch { $readyObj = $null }
                    }
                    if ($readyObj) { break }
                    if ([datetime]::UtcNow -ge $readyDeadline) { break }
                    $remainReady = [int][math]::Max(1, ($readyDeadline - [datetime]::UtcNow).TotalMilliseconds)
                    $readyTask = $proc.StandardOutput.ReadLineAsync()
                    if (-not $readyTask.Wait($remainReady)) { break }
                    $ready = [string]$readyTask.Result
                }
                if ($null -eq $readyObj) {
                    $err = ''
                    try { $err = [string]$Sync.PwWorkerErrTask.Result } catch { }
                    $Sync.PwWorker = $null
                    try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
                    Write-McpPayload (New-McpToolErrorJson "ProjectWise MCP worker failed to start: $err")
                    return
                }
                if (-not [bool]$readyObj.ok) {
                    $Sync.PwWorker = $null
                    try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
                    Write-McpPayload (New-McpToolErrorJson ("ProjectWise MCP worker init failed: " + [string]$readyObj.error))
                    return
                }
                $Sync.PwWorker = $proc
            }
            $proc.StandardInput.WriteLine($ReqJson)
            $proc.StandardInput.Flush()
            $deadline = [datetime]::UtcNow.AddMilliseconds($TimeoutMs)
            $resp = $null
            while ([datetime]::UtcNow -lt $deadline) {
                $remain = [int][math]::Max(1, ($deadline - [datetime]::UtcNow).TotalMilliseconds)
                $task = $proc.StandardOutput.ReadLineAsync()
                if (-not $task.Wait($remain)) { break }
                $line = [string]$task.Result
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if (-not $line.Trim().StartsWith('{')) { continue }
                try { $obj = $line | ConvertFrom-Json } catch { continue }
                if ($obj.id -and ([string]$obj.id -ne $RequestId)) { continue }
                $resp = $obj
                break
            }
            if ($null -eq $resp) {
                try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
                $Sync.PwWorker = $null
                Write-McpPayload (New-McpToolErrorJson 'ProjectWise tool timed out after 90s. Retry warm_projectwise_session, then the PW tool.')
                return
            }
            $idJson = $McpId | ConvertTo-Json -Compress
            if ($resp.ok) {
                $dataJson = if ($null -eq $resp.data) { '{}' } else { $resp.data | ConvertTo-Json -Depth 80 -Compress }
                if ([string]::IsNullOrWhiteSpace($dataJson)) { $dataJson = '{}' }
                Write-McpPayload ('{"jsonrpc":"2.0","id":' + $idJson + ',"result":{"content":[{"type":"text","text":' + ($dataJson | ConvertTo-Json -Compress) + '}],"structuredContent":' + $dataJson + ',"isError":false}}')
            } else {
                Write-McpPayload (New-McpToolErrorJson ([string]$resp.error))
            }
        } catch {
            Write-McpPayload (New-McpToolErrorJson $_.Exception.Message)
        } finally {
            $Sync.PwBusy = $false
        }
    }).AddArgument($script:McpSync).AddArgument($req).AddArgument($requestId).AddArgument($script:McpStdout).AddArgument($script:McpWriteLock).AddArgument($script:McpUtf8).AddArgument($McpId).AddArgument(90000).AddArgument((Join-Path $PSScriptRoot 'pw_qc_worker.ps1')).AddArgument($repoRoot).AddArgument($envMap)
    if (-not $script:PwWaiters) { $script:PwWaiters = New-Object System.Collections.ArrayList }
    [void]$script:PwWaiters.Add(@{ powershell = $ps; runspace = $rs })
    [void]$ps.BeginInvoke()
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
    $noLookupTools = @('get_recent_errors', 'get_process_health', 'get_audit_scan_history', 'warm_projectwise_session')
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
                $clientVersion = [string]$params.protocolVersion
                $supported = @('2024-11-05', '2025-03-26', '2025-06-18', '2025-11-25')
                $protocolVersion = if ($supported -contains $clientVersion) { $clientVersion } else { '2024-11-05' }
                # Literal JSON: Windows PowerShell ConvertTo-Json turns empty @{} into [] .
                Send-McpJsonResult -Id $id -ResultJson "{`"protocolVersion`":`"$protocolVersion`",`"capabilities`":{`"tools`":{}},`"serverInfo`":{`"name`":`"pw-qc-debug`",`"version`":`"1.0.0`"}}"
            }
            'tools/list' {
                Send-McpJsonResult -Id $id -ResultJson "{`"tools`":$script:ToolsListJson}"
            }
            'resources/list' {
                Send-McpJsonResult -Id $id -ResultJson $script:EmptyResourcesJson
            }
            'prompts/list' {
                Send-McpJsonResult -Id $id -ResultJson $script:EmptyPromptsJson
            }
            'tools/call' {
                $toolName = [string]$params.name
                $toolArgs = @{}
                if ($params.ContainsKey('arguments') -and $params.arguments) {
                    $toolArgs = ConvertTo-McpHashtable $params.arguments
                }
                if ($script:PwToolNames -contains $toolName) {
                    Start-McpPwToolAsync -McpId $id -ToolName $toolName -ToolArgs $toolArgs
                    break
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

Stop-McpPwWorker
