# PowerShell tool backend for pw-qc-debug MCP (invoked by server.py).
# Worker mode: JSON lines on stdin/stdout. One-shot: -ToolName + -ArgumentsJson.

param(
    [string]$ToolName = '',
    [string]$ArgumentsJson = '{}',
    [string]$ArgumentsFile = '',
    [switch]$Worker
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'

$repoRoot = $env:PWQC_REPO_ROOT
if ([string]::IsNullOrWhiteSpace($repoRoot)) {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$modulesRoot = Join-Path $repoRoot 'modules'
$scriptsRoot = Join-Path $repoRoot 'scripts'
$script:ContextReady = $false

function Initialize-WorkerContext {
    if ($script:ContextReady) { return }
    # Keep bootstrap chatter off the JSON stdout pipe used by server.py.
    $prevInfo = $InformationPreference
    $prevVerbose = $VerbosePreference
    $InformationPreference = 'SilentlyContinue'
    $VerbosePreference = 'SilentlyContinue'
    try {
        . (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
        Import-QCModuleBootstrapSet -FeatureModules @(
            'Diagnostics\QC.DebugMcp.psm1'
            'Core\Core.Telemetry.psm1'
        ) -RequiredCommands @(
            'Initialize-QCDebugMcpContext'
            'Get-QCAppSettingsConfig'
            'Search-QCDebugSheet'
        ) -Context 'pw-qc-mcp worker'
        Import-QCModuleGlobal -RelativePath 'ProjectWise\PW.Connection.psm1'
        Import-QCModuleGlobal -RelativePath 'ProjectWise\PW.Discovery.psm1'
        Test-QCRequiredCommands -Names @('Invoke-PWAuthenticatedCommand') -Context 'pw-qc-mcp worker ProjectWise'
        $appSettings = $env:PWQC_APPSETTINGS
        if ([string]::IsNullOrWhiteSpace($appSettings)) {
            $appSettings = Join-Path $repoRoot 'appsettings.json'
        }
        Initialize-QCDebugMcpContext -AppSettingsPath $appSettings | Out-Null
        $script:ContextReady = $true
    } finally {
        $InformationPreference = $prevInfo
        $VerbosePreference = $prevVerbose
    }
}

function ConvertTo-WorkerHashtable {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($key in $Value.Keys) { $h[$key] = ConvertTo-WorkerHashtable $Value[$key] }
        return $h
    }
    if ($Value -is [string] -or $Value -is [System.ValueType] -or $Value -is [decimal]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $arr = @()
        foreach ($item in $Value) { $arr += ,(ConvertTo-WorkerHashtable $item) }
        return $arr
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-WorkerHashtable $p.Value }
        return $h
    }
    return $Value
}

$script:ToolDispatch = @{
    search_sheet                      = { param($a) Search-QCDebugSheet @a }
    get_sheet_identity                = { param($a) Get-QCDebugSheetIdentity @a }
    get_sheet_package_members         = { param($a) Get-QCDebugSheetPackageMembers @a }
    get_sheet_debug_timeline          = { param($a) Get-QCDebugSheetTimeline @a }
    get_notification_diagnostics      = { param($a) Get-QCDebugNotificationDiagnostics @a }
    get_data_integrity_report         = { param($a) Get-QCDebugDataIntegrityReport @a }
    get_qc_process_type_diagnostics   = { param($a) Get-QCDebugQcProcessTypeDiagnostics @a }
    compare_projectwise_to_database   = { param($a) Compare-QCProjectWiseToDatabase @a }
    get_recent_errors                 = { param($a) Get-QCDebugRecentErrors @a }
    get_process_health                = { param($a) Get-QCDebugProcessHealth @a }
    get_audit_scan_history            = { param($a) Get-QCDebugAuditScanHistory @a }
    get_job_timeline                  = { param($a) Get-QCDebugJobTimeline @a }
    get_document_debug_events         = { param($a) Get-QCDebugDocumentAutomationEvents @a }
    get_package_debug_events          = { param($a) Get-QCDebugPackageAutomationEvents @a }
    warm_projectwise_session          = { param($a) Invoke-McpWarmProjectWiseSession @a }
}

function Invoke-McpWarmProjectWiseSession {
    Initialize-WorkerContext
    $appSettings = $env:PWQC_APPSETTINGS
    if ([string]::IsNullOrWhiteSpace($appSettings)) {
        $appSettings = Join-Path $repoRoot 'appsettings.json'
    }
    $cfg = Get-QCAppSettingsConfig -Path $appSettings
    $ds = [string]$cfg.projectWise.datasourceName
    $credPath = [string]$cfg.projectWise.credentialPath
    if ([string]::IsNullOrWhiteSpace($ds) -or [string]::IsNullOrWhiteSpace($credPath)) {
        throw 'projectWise.datasourceName or credentialPath missing in appsettings.'
    }
    # Local scriptblock (not remoting): do not use $using:; close over a local copy.
    $connectedDs = $ds
    $result = Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -KeepSession -ScriptBlock {
        return @{ datasourceName = $connectedDs; connected = $true }
    }
    return @{
        source_tables = @()
        query_assumptions = @('Opened a persistent ProjectWise session for subsequent MCP tool calls.')
        data = $result
        warnings = @()
    }
}

function Convert-ToolArguments {
    param([hashtable]$Arguments)
    $map = @{
        sheet_number         = 'SheetNumber'
        document_guid        = 'DocumentGuid'
        package_id           = 'PackageId'
        document_path        = 'DocumentPath'
        sheet_name           = 'SheetName'
        job_id               = 'JobId'
        limit                = 'Limit'
        hours                = 'Hours'
        force_jsonl_fallback = 'ForceJsonlFallback'
    }
    $out = @{}
    foreach ($k in $Arguments.Keys) {
        $target = if ($map.ContainsKey($k)) { $map[$k] } else { $k }
        $out[$target] = $Arguments[$k]
    }
    return $out
}

function Get-BoundToolArguments {
    param(
        [hashtable]$Arguments,
        [string]$ToolName
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
    $noLookup = @('get_recent_errors', 'get_process_health', 'get_audit_scan_history', 'warm_projectwise_session')
    $requireLookup = $noLookup -notcontains $ToolName
    if ($requireLookup -and $ToolName -ne 'get_job_timeline') {
        if ($args.Keys.Count -eq 0 -or -not ($args.ContainsKey('sheet_number') -or $args.ContainsKey('document_guid') -or $args.ContainsKey('package_id') -or $args.ContainsKey('document_path') -or $args.ContainsKey('sheet_name') -or $args.ContainsKey('job_id'))) {
            throw 'Provide one of: sheet_number, document_guid, package_id, document_path, sheet_name, or job_id.'
        }
    }
    if ($ToolName -eq 'get_job_timeline') {
        if (-not $args.ContainsKey('job_id') -and -not ($args.ContainsKey('sheet_number') -or $args.ContainsKey('document_guid') -or $args.ContainsKey('package_id') -or $args.ContainsKey('document_path') -or $args.ContainsKey('sheet_name'))) {
            throw 'get_job_timeline requires job_id or a sheet/document/package lookup.'
        }
    }
    if ($args.ContainsKey('limit')) { $args['limit'] = [int]$args['limit'] }
    if ($args.ContainsKey('hours')) { $args['hours'] = [int]$args['hours'] }
    return $args
}

function Invoke-QcDebugTool {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Arguments = @{}
    )
    Initialize-WorkerContext
    if (-not $script:ToolDispatch.ContainsKey($Name)) {
        throw "Unknown tool: $Name"
    }
    $bound = Get-BoundToolArguments -Arguments $Arguments -ToolName $Name
    return & $script:ToolDispatch[$Name] (Convert-ToolArguments -Arguments $bound)
}

function Write-WorkerResponse {
    param(
        [bool]$Ok,
        $Data = $null,
        [string]$Error = '',
        [string]$Id = '',
        [string]$Tool = ''
    )
    # Correlation fields (id/tool) let server.py discard stale stdout lines after
    # cancelled or overlapped MCP tool calls — without them, responses desync and
    # callers receive an unrelated prior payload (often sheet-identity).
    $payload = if ($Ok) { @{ ok = $true; data = $Data } } else { @{ ok = $false; error = $Error } }
    if (-not [string]::IsNullOrWhiteSpace($Id)) { $payload['id'] = $Id }
    if (-not [string]::IsNullOrWhiteSpace($Tool)) { $payload['tool'] = $Tool }
    $json = $payload | ConvertTo-Json -Depth 80 -Compress
    if ($script:WorkerJsonOut) {
        $script:WorkerJsonOut.WriteLine($json)
        $script:WorkerJsonOut.Flush()
    } else {
        [Console]::Out.WriteLine($json)
        [Console]::Out.Flush()
    }
}

function Enable-WorkerStdoutGuard {
    if ($script:WorkerStdoutGuard) { return }
    $script:WorkerJsonOut = [Console]::Out
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $errWriter = New-Object System.IO.StreamWriter([Console]::OpenStandardError(), $utf8)
    $errWriter.AutoFlush = $true
    [Console]::SetOut($errWriter)
    $script:WorkerStdoutGuard = $true
}

function Handle-WorkerRequest {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $requestId = ''
    $tool = ''
    try {
        $req = ConvertTo-WorkerHashtable ($Line | ConvertFrom-Json)
        if ($req.ContainsKey('id') -and $null -ne $req.id) { $requestId = [string]$req.id }
        $tool = [string]$req.tool
        $arguments = @{}
        if ($req.arguments) {
            $arguments = ConvertTo-WorkerHashtable $req.arguments
        }
        $data = Invoke-QcDebugTool -Name $tool -Arguments $arguments
        Write-WorkerResponse -Ok $true -Data $data -Id $requestId -Tool $tool
    } catch {
        Write-WorkerResponse -Ok $false -Error $_.Exception.Message -Id $requestId -Tool $tool
    }
}

if ($Worker) {
    try {
        Enable-WorkerStdoutGuard
        Initialize-WorkerContext
        Write-WorkerResponse -Ok $true -Data @{ ready = $true } -Tool 'worker_ready'
        while ($true) {
            $line = [Console]::In.ReadLine()
            if ($null -eq $line) { break }
            Handle-WorkerRequest -Line $line
        }
    } catch {
        Write-WorkerResponse -Ok $false -Error $_.Exception.Message -Tool 'worker_ready'
        exit 1
    }
    exit 0
}

try {
    Enable-WorkerStdoutGuard
    if (-not [string]::IsNullOrWhiteSpace($ArgumentsFile)) {
        if (-not (Test-Path -LiteralPath $ArgumentsFile)) {
            throw "ArgumentsFile not found: $ArgumentsFile"
        }
        $ArgumentsJson = Get-Content -LiteralPath $ArgumentsFile -Raw
    }
    $arguments = ConvertTo-WorkerHashtable ($ArgumentsJson | ConvertFrom-Json)
    if (-not $arguments) { $arguments = @{} }
    $data = Invoke-QcDebugTool -Name $ToolName -Arguments $arguments
    Write-WorkerResponse -Ok $true -Data $data -Tool $ToolName
} catch {
    Write-WorkerResponse -Ok $false -Error $_.Exception.Message -Tool $ToolName
    exit 1
}
