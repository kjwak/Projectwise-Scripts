$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$env:PWQC_APPSETTINGS = Join-Path $repoRoot 'appsettings.json'
$env:PWQC_SQL_SERVER = '192.168.22.90'
$env:PWQC_SQL_DATABASE = 'QC_Pipeline'
$env:PWQC_SQL_TRUST_CERT = 'yes'

$init = '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'
$bytes = [System.Text.Encoding]::UTF8.GetBytes($init)
$frame = "Content-Length: $($bytes.Length)`r`n`r`n$init"

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $PSScriptRoot 'server.ps1')`""
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true

$p = [System.Diagnostics.Process]::Start($psi)
$p.StandardInput.Write($frame)
$list = '{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
$listBytes = [System.Text.Encoding]::UTF8.GetBytes($list)
$p.StandardInput.Write("Content-Length: $($listBytes.Length)`r`n`r`n$list")
$p.StandardInput.Close()
Start-Sleep -Seconds 5
$out = $p.StandardOutput.ReadToEnd()
$err = $p.StandardError.ReadToEnd()
if (-not $p.HasExited) { $p.Kill() }

if ($out -notmatch '"protocolVersion"') { throw "initialize response missing: $out" }
if ($out -notmatch 'search_sheet') { throw "tools/list missing search_sheet: $out" }
if ($err) { throw "unexpected stderr: $err" }
Write-Host 'stdio handshake: PASS' -ForegroundColor Green
