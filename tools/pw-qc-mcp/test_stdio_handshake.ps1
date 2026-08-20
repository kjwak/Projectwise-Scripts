$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$env:PWQC_APPSETTINGS = Join-Path $repoRoot 'appsettings.json'
$env:PWQC_SQL_SERVER = '192.168.22.90'
$env:PWQC_SQL_DATABASE = 'QC_Pipeline'
$env:PWQC_SQL_TRUST_CERT = 'yes'
$env:PWQC_REPO_ROOT = $repoRoot

$init = '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'
$initialized = '{"jsonrpc":"2.0","method":"notifications/initialized"}'
$list = '{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'

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
$outTask = $p.StandardOutput.ReadToEndAsync()
$errTask = $p.StandardError.ReadToEndAsync()
$p.StandardInput.Write("$init`n$initialized`n$list`n")
$p.StandardInput.Close()
if (-not $p.WaitForExit(10000)) {
    $p.Kill()
    throw 'server did not exit after stdin closed'
}
$out = $outTask.GetAwaiter().GetResult()
$err = $errTask.GetAwaiter().GetResult()

if ($out -match 'Content-Length:') { throw "unexpected Content-Length framing: $out" }
if ($out -match 'Windows PowerShell') { throw "PowerShell logo leaked to stdout: $out" }
if ($out -notmatch '"protocolVersion"') { throw "initialize response missing: $out" }
if ($out -notmatch 'search_sheet') { throw "tools/list missing search_sheet: $out`nstderr=$err" }
if ($err) { throw "unexpected stderr: $err" }
Write-Host 'stdio handshake: PASS' -ForegroundColor Green
