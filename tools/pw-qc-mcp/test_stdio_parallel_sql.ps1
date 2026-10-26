# Prove SQL MCP tools are not blocked behind ProjectWise Connect-PW.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$env:PWQC_APPSETTINGS = Join-Path $repoRoot 'appsettings.json'
$env:PWQC_SQL_SERVER = '192.168.22.90'
$env:PWQC_SQL_DATABASE = 'QC_Pipeline'
$env:PWQC_SQL_TRUST_CERT = 'yes'
$env:PWQC_REPO_ROOT = $repoRoot

$init = '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'
$initialized = '{"jsonrpc":"2.0","method":"notifications/initialized"}'
$warm = '{"jsonrpc":"2.0","id":10,"method":"tools/call","params":{"name":"warm_projectwise_session","arguments":{}}}'
$health = '{"jsonrpc":"2.0","id":11,"method":"tools/call","params":{"name":"get_process_health","arguments":{}}}'

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$psi.Arguments = "-NoLogo -MTA -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$(Join-Path $PSScriptRoot 'server.ps1')`""
$psi.WorkingDirectory = $repoRoot
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true

$p = [System.Diagnostics.Process]::Start($psi)
try {
    $p.StandardInput.Write("$init`n$initialized`n$warm`n$health`n")
    $p.StandardInput.Flush()

    $gotHealth = $false
    $healthMs = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt 15000) {
        $remain = [int][math]::Max(1, 15000 - $sw.ElapsedMilliseconds)
        $task = $p.StandardOutput.ReadLineAsync()
        if (-not $task.Wait($remain)) { break }
        $line = [string]$task.Result
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $msg = $line | ConvertFrom-Json
        if ([string]$msg.id -eq '11') {
            $gotHealth = $true
            $healthMs = $sw.ElapsedMilliseconds
            break
        }
    }

    if (-not $gotHealth) {
        throw "get_process_health (id 11) did not return within 15s while warm_projectwise_session was in flight. SQL tools are still blocked on PW."
    }
    Write-Host "stdio parallel SQL: PASS (get_process_health in ${healthMs}ms while PW warm dispatched)" -ForegroundColor Green
} finally {
    try { $p.StandardInput.Close() } catch { }
    if (-not $p.WaitForExit(8000)) {
        try { $p.Kill() } catch { }
    }
}
