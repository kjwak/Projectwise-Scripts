<# Root entrypoint: live pipeline + dashboard.
Prefer running: scripts\Start-QCPipelineDashboard.ps1
#>

$target = Join-Path $PSScriptRoot 'scripts\Start-QCPipelineDashboard.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Host "Dashboard script not found: $target" -ForegroundColor Red
    Write-Host "Run this from the repo root (folder that contains the scripts\ subfolder)." -ForegroundColor Yellow
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host ''
        Write-Host 'Press Enter to close...' -ForegroundColor Yellow
        try { $null = Read-Host } catch { }
    }
    exit 1
}

try {
    & $target @args
    exit $LASTEXITCODE
} catch {
    Write-Host ''
    Write-Host "Failed to start dashboard: $($_.Exception.Message)" -ForegroundColor Red
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host ''
        Write-Host 'Press Enter to close...' -ForegroundColor Yellow
        try { $null = Read-Host } catch { }
    }
    exit 1
}
