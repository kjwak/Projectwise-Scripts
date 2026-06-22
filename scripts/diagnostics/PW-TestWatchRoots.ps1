<#
.SYNOPSIS
Quick diagnostic for ProjectWise watchList root expansion.
#>

[CmdletBinding()]
param(
    [string]$AppSettingsPath = ''
)

$ErrorActionPreference = 'Stop'

function _Get-ThisScriptDir {
    try {
        if ($PSScriptRoot -and -not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { return $PSScriptRoot }
    } catch { }
    try {
        $p = $MyInvocation.MyCommand.Path
        if ($p -and (Test-Path -LiteralPath $p)) { return (Split-Path -Parent $p) }
    } catch { }
    return (Get-Location).Path
}

$scriptDir = _Get-ThisScriptDir
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = (Join-Path $repoRoot 'appsettings.json')
}

Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$pw = ConvertTo-HashtableDeep -Value $config.projectWise
$ds = [string]$pw.datasourceName
$credRes = Get-PWCredentialFromFile -CredentialPath ([string]$pw.credentialPath)
if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
$conn = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $conn.IsSuccess) { throw ($conn.Code + ': ' + $conn.Message) }

try {
    $watchList = ConvertTo-HashtableDeep -Value $pw.watchList
    foreach ($r in @($watchList.roots)) {
        $rh = ConvertTo-HashtableDeep -Value $r
        $rootPath = [string]$rh.path
        $suffix = if ($rh.sheetsPathFromProject) { [string]$rh.sheetsPathFromProject } else { 'CADD\Sheets' }
        $depth = if ($rh.projectDepth) { [int]$rh.projectDepth } else { 1 }
        $found = @(Find-PWSheetsFoldersUnderRoot -RootPath $rootPath -SheetsSuffix $suffix -DatasourceName $ds -ProjectDepth $depth)
        Write-Host ("Root={0} depth={1} -> sheetsFolders={2}" -f $rootPath, $depth, $found.Count) -ForegroundColor Cyan
        $found | Select-Object -First 10 FolderPath,OneLevelDeep | Format-Table -AutoSize
        foreach ($x in @($found | Select-Object -First 3)) {
            $fp = if ($x -and ($x -is [hashtable]) -and $x.ContainsKey('FolderPath')) { [string]$x.FolderPath } else { $null }
            $len = $(if ($fp) { $fp.Length } else { 0 })
            Write-Host ("  sample FolderPath='{0}' len={1}" -f $fp, $len) -ForegroundColor DarkGray
        }
        Write-Host ''
    }
} finally {
    try { Disconnect-PW | Out-Null } catch { }
}

