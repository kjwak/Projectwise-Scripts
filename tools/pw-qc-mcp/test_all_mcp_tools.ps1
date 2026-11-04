# Smoke-test all pw-qc-debug MCP tools via pw_qc_worker.ps1 (one-shot mode).
param(
    [string]$SheetNumber = '080J082001ca001',
    [switch]$SkipProjectWise
)

$ErrorActionPreference = 'Stop'
$repoRoot = if ($env:PWQC_REPO_ROOT) { $env:PWQC_REPO_ROOT } else { Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$worker = Join-Path $PSScriptRoot 'pw_qc_worker.ps1'
$env:PWQC_REPO_ROOT = $repoRoot
if (-not $env:PWQC_APPSETTINGS) {
    $env:PWQC_APPSETTINGS = Join-Path $repoRoot 'appsettings.json'
}

function Invoke-WorkerTool {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Arguments = @{}
    )
    $argsFile = Join-Path ([System.IO.Path]::GetTempPath()) ("pwqc-mcp-args-{0}.json" -f [guid]::NewGuid().ToString('N'))
    try {
        ($Arguments | ConvertTo-Json -Compress -Depth 5) | Set-Content -LiteralPath $argsFile -Encoding UTF8 -NoNewline
        $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $worker -ToolName $Name -ArgumentsFile $argsFile 2>&1
        $line = ($out | Select-Object -Last 1)
        $json = $line | ConvertFrom-Json
        if (-not $json.ok) {
            throw $json.error
        }
        return $json.data
    } finally {
        Remove-Item -LiteralPath $argsFile -Force -ErrorAction SilentlyContinue
    }
}

$tools = @(
    @{ Name = 'get_process_health'; Args = @{} }
    @{ Name = 'get_recent_errors'; Args = @{ limit = 5; hours = 24 } }
    @{ Name = 'get_audit_scan_history'; Args = @{ limit = 5; hours = 72 } }
    @{ Name = 'search_sheet'; Args = @{ sheet_number = $SheetNumber } }
    @{ Name = 'get_sheet_identity'; Args = @{ sheet_number = $SheetNumber } }
    @{ Name = 'get_sheet_package_members'; Args = @{ sheet_number = $SheetNumber } }
    @{ Name = 'get_sheet_debug_timeline'; Args = @{ sheet_number = $SheetNumber; limit = 5 } }
    @{ Name = 'get_notification_diagnostics'; Args = @{ sheet_number = $SheetNumber; limit = 5 } }
    @{ Name = 'get_data_integrity_report'; Args = @{ sheet_number = $SheetNumber } }
    @{ Name = 'get_qc_process_type_diagnostics'; Args = @{ sheet_number = $SheetNumber } }
    @{ Name = 'get_document_debug_events'; Args = @{ sheet_number = $SheetNumber; limit = 5 } }
    @{ Name = 'get_package_debug_events'; Args = @{ sheet_number = $SheetNumber; limit = 5 } }
    @{ Name = 'get_job_timeline'; Args = @{ sheet_number = $SheetNumber; limit = 5 } }
)
if (-not $SkipProjectWise.IsPresent) {
    $tools += @(
        @{ Name = 'warm_projectwise_session'; Args = @{} }
        @{ Name = 'compare_projectwise_to_database'; Args = @{ sheet_number = $SheetNumber } }
    )
}

$fail = 0
foreach ($tool in $tools) {
    Write-Host "=== $($tool.Name) ===" -ForegroundColor Cyan
    try {
        $null = Invoke-WorkerTool -Name $tool.Name -Arguments $tool.Args
        Write-Host 'OK' -ForegroundColor Green
    } catch {
        Write-Host "FAIL: $($_.Exception.Message)" -ForegroundColor Red
        $fail++
    }
}

if ($fail -gt 0) {
    Write-Host "test_all_mcp_tools.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_all_mcp_tools.ps1 passed' -ForegroundColor Green
exit 0
