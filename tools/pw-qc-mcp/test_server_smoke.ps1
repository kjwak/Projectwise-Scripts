# Smoke test for QC.DebugMcp diagnostics (direct function calls, no MCP transport).
param(
    [string]$SheetNumber = '080J082001ca001',
    [string]$DocumentPath = '',
    [string]$AppSettingsPath = ''
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Diagnostics\QC.DebugMcp.psm1') -Force
Initialize-QCDebugMcpContext -AppSettingsPath $AppSettingsPath | Out-Null

function Show-Result {
    param([string]$Name, [scriptblock]$Block)
    Write-Host "`n=== $Name ===" -ForegroundColor Cyan
    try {
        $result = & $Block
        $json = $result | ConvertTo-Json -Depth 20
        if ($json.Length -gt 6000) {
            Write-Host ($json.Substring(0, 6000))
            Write-Host '... [truncated]'
        } else {
            Write-Host $json
        }
        return $true
    } catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

$ok = $true
if (-not [string]::IsNullOrWhiteSpace($DocumentPath)) {
    $ok = (Show-Result 'search_sheet (document_path)' { Search-QCDebugSheet -DocumentPath $DocumentPath }) -and $ok
    $ok = (Show-Result 'get_sheet_identity (document_path)' { Get-QCDebugSheetIdentity -DocumentPath $DocumentPath }) -and $ok
} else {
    $ok = (Show-Result 'search_sheet' { Search-QCDebugSheet -SheetNumber $SheetNumber }) -and $ok
    $ok = (Show-Result 'get_sheet_identity' { Get-QCDebugSheetIdentity -SheetNumber $SheetNumber }) -and $ok
    $ok = (Show-Result 'get_sheet_debug_timeline' { Get-QCDebugSheetTimeline -SheetNumber $SheetNumber -Limit 10 }) -and $ok
    $ok = (Show-Result 'get_data_integrity_report' { Get-QCDebugDataIntegrityReport -SheetNumber $SheetNumber }) -and $ok
}

if (-not $ok) { exit 1 }
Write-Host "`nAll smoke tests passed." -ForegroundColor Green
