# Regression: worker request/response correlation (id + tool).
# Proves that overlapping/stale stdout lines cannot be mistaken for the current tool.
$ErrorActionPreference = 'Stop'
$repoRoot = if ($env:PWQC_REPO_ROOT) { $env:PWQC_REPO_ROOT } else { Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$worker = Join-Path $PSScriptRoot 'pw_qc_worker.ps1'
$env:PWQC_REPO_ROOT = $repoRoot
if (-not $env:PWQC_APPSETTINGS) { $env:PWQC_APPSETTINGS = Join-Path $repoRoot 'appsettings.json' }
if (-not $env:PWQC_SQL_SERVER) { $env:PWQC_SQL_SERVER = '192.168.22.90' }
if (-not $env:PWQC_SQL_DATABASE) { $env:PWQC_SQL_DATABASE = 'QC_Pipeline' }
if (-not $env:PWQC_SQL_TRUST_CERT) { $env:PWQC_SQL_TRUST_CERT = 'yes' }

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$psi.Arguments = "-MTA -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$worker`" -Worker"
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$psi.EnvironmentVariables['PWQC_REPO_ROOT'] = $env:PWQC_REPO_ROOT
$psi.EnvironmentVariables['PWQC_APPSETTINGS'] = $env:PWQC_APPSETTINGS
$psi.EnvironmentVariables['PWQC_SQL_SERVER'] = $env:PWQC_SQL_SERVER
$psi.EnvironmentVariables['PWQC_SQL_DATABASE'] = $env:PWQC_SQL_DATABASE
$psi.EnvironmentVariables['PWQC_SQL_TRUST_CERT'] = $env:PWQC_SQL_TRUST_CERT

$p = [System.Diagnostics.Process]::Start($psi)
try {
    $ready = $p.StandardOutput.ReadLine()
    $readyObj = $ready | ConvertFrom-Json
    if (-not $readyObj.ok) { throw "worker init failed: $ready" }

    $id1 = [guid]::NewGuid().ToString('N')
    $id2 = [guid]::NewGuid().ToString('N')
    $id3 = [guid]::NewGuid().ToString('N')

    $p.StandardInput.WriteLine((@{ id = $id1; tool = 'get_recent_errors'; arguments = @{ limit = 3; hours = 3 } } | ConvertTo-Json -Compress -Depth 5))
    $p.StandardInput.WriteLine((@{ id = $id2; tool = 'get_process_health'; arguments = @{} } | ConvertTo-Json -Compress -Depth 5))
    $p.StandardInput.WriteLine((@{ id = $id3; tool = 'get_audit_scan_history'; arguments = @{ limit = 3; hours = 3 } } | ConvertTo-Json -Compress -Depth 5))
    $p.StandardInput.Flush()

    $seen = @{}
    for ($i = 0; $i -lt 3; $i++) {
        $line = $p.StandardOutput.ReadLine()
        if (-not $line) { throw "worker closed stdout early (got $i responses)" }
        $obj = $line | ConvertFrom-Json
        if (-not $obj.id) { throw "response missing id: $line" }
        if (-not $obj.tool) { throw "response missing tool: $line" }
        if ($seen.ContainsKey([string]$obj.id)) { throw "duplicate id $($obj.id)" }
        $seen[[string]$obj.id] = [string]$obj.tool
        # Correlation is the contract under test. SQL may be unreachable from some
        # agent shells; still require id/tool echo and (on success) payload shape.
        if ($obj.ok) {
            switch ([string]$obj.tool) {
                'get_recent_errors' {
                    if (-not ($obj.data.PSObject.Properties.Name -contains 'events')) {
                        throw "get_recent_errors missing events payload: $line"
                    }
                }
                'get_process_health' {
                    if (-not ($obj.data.PSObject.Properties.Name -contains 'processes')) {
                        throw "get_process_health missing processes payload: $line"
                    }
                }
                'get_audit_scan_history' {
                    if (-not ($obj.data.PSObject.Properties.Name -contains 'events')) {
                        throw "get_audit_scan_history missing events payload: $line"
                    }
                }
                default { throw "unexpected tool $($obj.tool)" }
            }
        } else {
            Write-Host ("NOTE: $($obj.tool) returned error (correlation still OK): $($obj.error)") -ForegroundColor Yellow
        }
    }

    if ($seen[$id1] -ne 'get_recent_errors') { throw "id1 mapped to $($seen[$id1])" }
    if ($seen[$id2] -ne 'get_process_health') { throw "id2 mapped to $($seen[$id2])" }
    if ($seen[$id3] -ne 'get_audit_scan_history') { throw "id3 mapped to $($seen[$id3])" }

    Write-Host 'test_worker_correlation.ps1: PASS' -ForegroundColor Green
    exit 0
}
finally {
    try { $p.StandardInput.Close() } catch {}
    if (-not $p.HasExited) { $p.Kill() }
    $p.Dispose()
}
