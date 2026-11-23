# Run the Windows developer baseline with synthetic fixtures in a fresh directory.
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$python = (Get-Command python.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
$outputRoot = (New-Item -ItemType Directory -Path $OutputDirectory).FullName
$scratch = (New-Item -ItemType Directory -Path (Join-Path $outputRoot 'scratch')).FullName
$summary = Join-Path $outputRoot 'summary.txt'
$offlineConfig = Join-Path $scratch 'offline.json'
'{"database":{"enabled":false}}' | Set-Content -LiteralPath $offlineConfig -Encoding ASCII

function Write-BaselineSummary([string]$Line) {
    Write-Host $Line
    Add-Content -LiteralPath $summary -Value $Line
}

function Invoke-BaselineSuite([string]$Name, [string]$Executable, [string[]]$Arguments) {
    $log = Join-Path $outputRoot "$Name.log"
    Write-BaselineSummary ('> & "' + $Executable + '" "' + ($Arguments -join '" "') + '"')
    # Capture Windows PowerShell's redirected native stderr without stopping the matrix.
    $ErrorActionPreference = 'Continue'
    & $Executable @Arguments *> $log
    $code = $LASTEXITCODE
    $result = 'PASS'
    if ($code -ne 0) {
        $result = 'FAIL'
    } elseif (Select-String -LiteralPath $log -Pattern '^SKIP:' -Quiet) {
        $result = 'SKIP'
    }
    Write-BaselineSummary "$result $Name (exit $code, $Name.log)"
    return $result
}

$before = @{
    TEMP = $env:TEMP
    TMP = $env:TMP
    PYTHONDONTWRITEBYTECODE = $env:PYTHONDONTWRITEBYTECODE
    PYTEST_ADDOPTS = $env:PYTEST_ADDOPTS
    PYTEST_DISABLE_PLUGIN_AUTOLOAD = $env:PYTEST_DISABLE_PLUGIN_AUTOLOAD
}
$results = @()
Push-Location -LiteralPath $repoRoot
try {
    $env:TEMP = $scratch
    $env:TMP = $scratch
    $env:PYTHONDONTWRITEBYTECODE = '1'
    $env:PYTEST_ADDOPTS = ''
    $env:PYTEST_DISABLE_PLUGIN_AUTOLOAD = '1'

    $revision = & git rev-parse HEAD
    if ($LASTEXITCODE -ne 0) { throw 'Run the baseline from a Git checkout.' }
    Write-BaselineSummary "Revision: $revision"
    Write-BaselineSummary "Scratch: $scratch"
    $version = & $windowsPowerShell -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
    Write-BaselineSummary "Windows PowerShell: $version"
    $pwsh = Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pwsh) {
        $version = & $pwsh.Source -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
        Write-BaselineSummary "PowerShell: $version"
    }
    $versions = & $python -c 'import sys, pytest, fitz, pikepdf; print(sys.version.split()[0], pytest.__version__, fitz.VersionBind, pikepdf.__version__)'
    if ($LASTEXITCODE -ne 0) { throw 'Install tools/overlay/requirements-dev.txt in your developer Python environment.' }
    Write-BaselineSummary "Python / pytest / PyMuPDF / pikepdf: $versions"

    $results += Invoke-BaselineSuite -Name 'pytest' -Executable $python -Arguments @(
        '-m', 'pytest', '-q', '-p', 'no:cacheprovider', '--basetemp', (Join-Path $scratch 'pytest'), '-m', 'not project_fixture'
    )
    $suiteNames = @(
        'queue_json', 'queue_retention', 'config_json_comments', 'config_profile_merge', 'effective_dry_run_policy',
        'module_inventory', 'module_folder_shims', 'watcher_module_bootstrap',
        'watch_foundation_restore', 'statusset_publication', 'audit_events_db'
    )
    foreach ($suite in $suiteNames) {
        $testArguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ".\test\powershell\test_$suite.ps1")
        if ($suite -eq 'audit_events_db') {
            $testArguments += @('-AppSettingsPath', $offlineConfig)
        }
        $results += Invoke-BaselineSuite -Name $suite -Executable $windowsPowerShell -Arguments $testArguments
    }
    $passed = @($results -eq 'PASS').Count
    $failed = @($results -eq 'FAIL').Count
    $skipped = @($results -eq 'SKIP').Count
    Write-BaselineSummary "Suites: $passed passed, $failed failed, $skipped skipped. See pytest.log for Python case counts."
} finally {
    Pop-Location
    foreach ($name in $before.Keys) {
        [Environment]::SetEnvironmentVariable($name, $before[$name], 'Process')
    }
}

if ($failed -gt 0) { exit 1 }
