<#
.SYNOPSIS
Local smoke test for Microsoft Graph QC notifications.

.DESCRIPTION
Reads notifications.graph from appsettings.json (+ appsettings.local.json if present)
and sends a minimal message (subject and body: "test").

By default uses notifications.dryRun from config (no Graph API call when true).
Pass -Live to send a real email (overrides dryRun).

.EXAMPLE
.\scripts\Test-QCNotificationGraph.ps1 -To you@company.com

.EXAMPLE
.\scripts\Test-QCNotificationGraph.ps1 -To you@company.com -Live
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$To,

    [string[]]$Cc = @(),

    [string]$AppSettingsPath = '',

    [switch]$Live
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    throw @"
PSScriptRoot is empty. Run from the repo root, for example:

  .\scripts\Test-QCNotificationGraph.ps1 -To you@company.com -Live
"@
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}
if (-not (Test-Path -LiteralPath $AppSettingsPath)) {
    throw "appsettings not found: $AppSettingsPath"
}

function _Test-RemoveJsonComments([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $sb = New-Object System.Text.StringBuilder
    $len = $Text.Length
    $inString = $false
    $escape = $false
    $i = 0
    while ($i -lt $len) {
        $c = $Text[$i]
        if ($inString) {
            [void]$sb.Append($c)
            if ($escape) { $escape = $false }
            elseif ($c -eq '\') { $escape = $true }
            elseif ($c -eq '"') { $inString = $false }
            $i++
            continue
        }
        if ($c -eq '"') {
            $inString = $true
            [void]$sb.Append($c)
            $i++
            continue
        }
        if ($c -eq '/' -and ($i + 1) -lt $len) {
            $n = $Text[$i + 1]
            if ($n -eq '/') {
                $i += 2
                while ($i -lt $len -and $Text[$i] -ne "`n" -and $Text[$i] -ne "`r") { $i++ }
                continue
            }
            if ($n -eq '*') {
                $i += 2
                while ($i + 1 -lt $len -and -not ($Text[$i] -eq '*' -and $Text[$i + 1] -eq '/')) { $i++ }
                $i += 2
                continue
            }
        }
        [void]$sb.Append($c)
        $i++
    }
    return $sb.ToString()
}

function _Test-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.ValueType]) { return $Value }
    if ($Value -is [hashtable] -or $Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[$k] = _Test-ToHashtable $Value[$k] }
        return $h
    }
    if ($Value -is [System.Array] -or ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string]))) {
        $list = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$list.Add((_Test-ToHashtable $item)) }
        return @($list.ToArray())
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = _Test-ToHashtable $p.Value }
        return $h
    }
    return $Value
}

function _Test-MergeHashtable([hashtable]$Base, [hashtable]$Overlay) {
    if ($null -eq $Overlay) { return $Base }
    if ($null -eq $Base) { return $Overlay }
    $merged = @{}
    foreach ($key in $Base.Keys) { $merged[$key] = $Base[$key] }
    foreach ($key in $Overlay.Keys) {
        if ($merged.ContainsKey($key) -and $merged[$key] -is [hashtable] -and $Overlay[$key] -is [hashtable]) {
            $merged[$key] = _Test-MergeHashtable $merged[$key] $Overlay[$key]
        } else {
            $merged[$key] = $Overlay[$key]
        }
    }
    return $merged
}

function _Test-ReadAppSettingsLayer([string]$Path) {
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $json = _Test-RemoveJsonComments -Text $raw
    return _Test-ToHashtable ($json | ConvertFrom-Json -ErrorAction Stop)
}

function Get-TestNotificationAppSettings([string]$Path) {
    $dir = [System.IO.Path]::GetDirectoryName((Resolve-Path -LiteralPath $Path).Path)
    $name = [System.IO.Path]::GetFileName($Path)
    $chain = @($Path)
    if ($name -eq 'appsettings.json') {
        $local = Join-Path $dir 'appsettings.local.json'
        if (Test-Path -LiteralPath $local) {
            $chain += (Resolve-Path -LiteralPath $local).Path
        }
        $secrets = Join-Path $dir 'appsettings.secrets.json'
        if (Test-Path -LiteralPath $secrets) {
            $chain += (Resolve-Path -LiteralPath $secrets).Path
        }
    }

    $cfg = $null
    foreach ($filePath in $chain) {
        $layer = _Test-ReadAppSettingsLayer -Path $filePath
        if ($null -eq $cfg) { $cfg = $layer } else { $cfg = _Test-MergeHashtable $cfg $layer }
    }
    if ($null -eq $cfg) { $cfg = @{} }
    return $cfg
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Notifications\QC.NotificationGraph.psm1'
) -RequiredCommands @(
    'Send-QCNotificationGraph'
    'Test-QCNotificationGraphConfigured'
) -Context 'Test-QCNotificationGraph bootstrap'

$config = Get-TestNotificationAppSettings -Path $AppSettingsPath

$notifications = @{}
if ($config.ContainsKey('notifications') -and $config.notifications) {
    $norm = _Test-ToHashtable $config.notifications
    if ($norm) { $notifications = $norm }
}

$graph = @{}
if ($notifications.ContainsKey('graph') -and $notifications.graph) {
    $g = _Test-ToHashtable $notifications.graph
    if ($g) { $graph = $g }
}

$validation = Test-QCNotificationGraphConfigured -GraphSettings $graph
if (-not $validation.configured) {
    throw ('Graph not configured. Missing: ' + ($validation.missing -join ', ') + '. Set notifications.graph in ' + $AppSettingsPath)
}

$dryRun = $true
if ($notifications.ContainsKey('dryRun')) {
    $dryRun = [bool]$notifications.dryRun
}
if ($Live) { $dryRun = $false }

$payload = @{
    eventType = 'GRAPH_TEST'
    project = 'local-test'
    documentName = 'test.pdf'
    documentPath = 'local-test/test.pdf'
    documentGuid = '00000000-0000-0000-0000-000000000099'
    previousState = 'Test'
    currentState = 'Test'
    actionRequired = 'Graph smoke test.'
    sourceJobId = 'graph-test'
    subject = 'test'
    body = 'test'
    to = @($To.Trim())
    cc = @($Cc | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    reviewers = @()
    designers = @()
}

Write-Host "Repo root:   $repoRoot" -ForegroundColor Cyan
Write-Host "AppSettings: $AppSettingsPath" -ForegroundColor Cyan
Write-Host "Sender:      $($graph.senderMailbox)" -ForegroundColor Cyan
Write-Host "To:          $($payload.to -join ', ')" -ForegroundColor Cyan
if ($payload.cc.Count -gt 0) {
    Write-Host "Cc:          $($payload.cc -join ', ')" -ForegroundColor Cyan
}
Write-Host "DryRun:      $dryRun  (pass -Live to send mail)" -ForegroundColor Cyan

$result = Send-QCNotificationGraph -GraphSettings $graph -Payload $payload -DryRun:$dryRun

$out = @{
    success = $result.IsSuccess
    code = $result.Code
    message = $result.Message
    data = $result.Data
}
Write-Host ($out | ConvertTo-Json -Depth 12)

if (-not $result.IsSuccess) {
    throw $result.Message
}

if ($dryRun) {
    Write-Host 'Dry run OK. Re-run with -Live to send a real email.' -ForegroundColor Green
} else {
    Write-Host 'Graph sendMail completed.' -ForegroundColor Green
}
