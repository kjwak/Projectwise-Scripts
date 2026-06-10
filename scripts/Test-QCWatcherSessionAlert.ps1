<#
.SYNOPSIS
Send a sample QC watcher ProjectWise session-lost alert email.

.EXAMPLE
.\scripts\Test-QCWatcherSessionAlert.ps1 -Live

.EXAMPLE
.\scripts\Test-QCWatcherSessionAlert.ps1 -To jflint@aztec.us -Live
#>
[CmdletBinding()]
param(
    [string]$To = '',
    [string]$AppSettingsPath = '',
    [switch]$Live,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

function _TWSA-RemoveJsonComments([string]$Text) {
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
        if ($c -eq '"') { $inString = $true; [void]$sb.Append($c); $i++; continue }
        if ($c -eq '/' -and ($i + 1) -lt $len) {
            $n = $Text[$i + 1]
            if ($n -eq '/') { $i += 2; while ($i -lt $len -and $Text[$i] -ne "`n" -and $Text[$i] -ne "`r") { $i++ }; continue }
            if ($n -eq '*') { $i += 2; while ($i + 1 -lt $len -and -not ($Text[$i] -eq '*' -and $Text[$i + 1] -eq '/')) { $i++ }; $i += 2; continue }
        }
        [void]$sb.Append($c)
        $i++
    }
    return $sb.ToString()
}

function _TWSA-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.ValueType]) { return $Value }
    if ($Value -is [hashtable] -or $Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[$k] = _TWSA-ToHashtable $Value[$k] }
        return $h
    }
    if ($Value -is [System.Array] -or ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string]))) {
        $list = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$list.Add((_TWSA-ToHashtable $item)) }
        return @($list.ToArray())
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = _TWSA-ToHashtable $p.Value }
        return $h
    }
    return $Value
}

function _TWSA-MergeHashtable([hashtable]$Base, [hashtable]$Overlay) {
    if ($null -eq $Overlay) { return $Base }
    if ($null -eq $Base) { return $Overlay }
    $merged = @{}
    foreach ($key in $Base.Keys) { $merged[$key] = $Base[$key] }
    foreach ($key in $Overlay.Keys) {
        if ($merged.ContainsKey($key) -and $merged[$key] -is [hashtable] -and $Overlay[$key] -is [hashtable]) {
            $merged[$key] = _TWSA-MergeHashtable $merged[$key] $Overlay[$key]
        } else {
            $merged[$key] = $Overlay[$key]
        }
    }
    return $merged
}

function _TWSA-ReadAppSettings([string]$Path) {
    $dir = [System.IO.Path]::GetDirectoryName((Resolve-Path -LiteralPath $Path).Path)
    $chain = @((Resolve-Path -LiteralPath $Path).Path)
    foreach ($extra in @('appsettings.local.json', 'appsettings.secrets.json')) {
        $p = Join-Path $dir $extra
        if (Test-Path -LiteralPath $p) { $chain += (Resolve-Path -LiteralPath $p).Path }
    }
    $cfg = $null
    foreach ($filePath in $chain) {
        $raw = Get-Content -LiteralPath $filePath -Raw -ErrorAction Stop
        $layer = _TWSA-ToHashtable ((_TWSA-RemoveJsonComments -Text $raw) | ConvertFrom-Json -ErrorAction Stop)
        if ($null -eq $cfg) { $cfg = $layer } else { $cfg = _TWSA-MergeHashtable $cfg $layer }
    }
    return $cfg
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/QC.NotificationGraph.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules/QC.WatcherAlerts.psm1') -Force

$config = _TWSA-ReadAppSettings -Path $AppSettingsPath
$alertSettings = Get-QCWatcherSessionAlertSettings -Config $config

if (-not [string]::IsNullOrWhiteSpace($To)) {
    $config.watcher = if ($config.watcher) { _TWSA-ToHashtable $config.watcher } else { @{} }
    if (-not $config.watcher) { $config.watcher = @{} }
    $config.watcher['sessionAlerts'] = @{
        enabled = $true
        recipients = @($To.Trim())
        importance = 'high'
        dedupeMinutes = $alertSettings.dedupeMinutes
    }
}

if ($Live) {
    if ($config.ContainsKey('notifications')) {
        $n = _TWSA-ToHashtable $config.notifications
        if ($n) {
            $n['dryRun'] = $false
            $config['notifications'] = $n
        }
    }
    $config['dryRun'] = $false
}

$ds = ''
if ($config.projectWise -and $config.projectWise.datasourceName) {
    $ds = [string]$config.projectWise.datasourceName
}

$details = @{
    detectedUtc = (Get-Date).ToUniversalTime().ToString('o')
    reason = 'manual_test'
    datasourceName = $ds
    tick = 'TEST'
    probeFolderPath = if ($alertSettings.probeFolderPath) { [string]$alertSettings.probeFolderPath } else { 'Documents\AZDOT 2024' }
    maxPwActTime = (Get-Date).AddMinutes(-15).ToString('yyyy-MM-dd HH:mm:ss')
    watermarkAfter = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    errorMessage = 'Manual preview send from Test-QCWatcherSessionAlert.ps1 (not a real outage).'
}

Write-Host "Recipients: $($alertSettings.recipients -join ', ')" -ForegroundColor Cyan
Write-Host "Importance: $($alertSettings.importance)" -ForegroundColor Cyan
Write-Host "Live send:  $([bool]$Live)" -ForegroundColor Cyan

$params = @{
    Config = $config
    Details = $details
}
if ($Force) { $params['Force'] = $true }

$result = Send-QCWatcherSessionLostAlert @params
Write-Host ($result | ConvertTo-Json -Depth 8)

if (-not $result.IsSuccess) { throw $result.Message }
if (-not $Live) {
    Write-Host 'Dry run / mock OK. Re-run with -Live to send a real email.' -ForegroundColor Green
} else {
    Write-Host 'Session-lost preview email sent.' -ForegroundColor Green
}
