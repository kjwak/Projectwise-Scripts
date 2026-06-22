<#
.SYNOPSIS
Returns ProjectWise workflow state counts for a folder tree (read-only).

.DESCRIPTION
Given a ProjectWise folder path (supports pw:\datasource\Documents\... form), connects to the
configured datasource and runs Get-PWFolderTreeDocumentStateCount (pwps_dab) for that folder.

This script is intentionally read-only.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\Get-PWFolderStateCounts.ps1 `
  -FolderPath "pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\Caltrans\...\CADD\Sheets\04-Layout\"
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification = 'Script accepts a file path to a credential file; no password is passed as plaintext.')]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [string]$DatasourceName,

    [Parameter(Mandatory = $false)]
    [string]$PwCredFilePath,

    [Parameter(Mandatory = $true)]
    [string]$FolderPath,

    [Parameter(Mandatory = $false)]
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'

function ConvertFrom-PWUriToFolderPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)

    $s = ($Value -as [string]).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s = $s.Trim().Trim('"').Trim("'").TrimEnd('\')
    $s = $s -replace '/', '\'

    # pw:\\server\Documents\foo\bar  => Documents\foo\bar
    if ($s -match '^(?i)pw:\\\\[^\\]+\\(.+)$') {
        $s = $Matches[1]
    }

    # Ensure canonical "Documents\..." prefix for pwps_dab compatibility.
    if ($s -notmatch '^(?i)Documents\\') {
        $s = ('Documents\' + $s.TrimStart('\'))
    }

    return $s.TrimEnd('\')
}

function _FailJson([string]$Code, [string]$Message, [hashtable]$Data = $null) {
    $out = [ordered]@{
        ok = $false
        code = $Code
        message = $Message
        data = $Data
    }
    if ($Pretty) { return ($out | ConvertTo-Json -Depth 7) }
    return ($out | ConvertTo-Json -Depth 7 -Compress)
}

$scriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptPath)) { throw 'Cannot resolve script path via $MyInvocation.MyCommand.Path.' }
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptPath))

if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Config.psm1'
    'ProjectWise\PW.Connection.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Get-QCTimestamp'
) -Context 'Get-PWFolderStateCounts bootstrap'

# Resolve config defaults from appsettings.json unless overridden.
$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$cfg = $cfgRes.Data.config
$valRes = Test-AppSettings -Config $cfg
if (-not $valRes.IsSuccess) { throw $valRes.Message }

if ([string]::IsNullOrWhiteSpace($DatasourceName)) {
    $DatasourceName = [string]$cfg['projectWise']['datasourceName']
}
if ([string]::IsNullOrWhiteSpace($PwCredFilePath)) {
    $PwCredFilePath = [string]$cfg['projectWise']['credentialPath']
}
if ([string]::IsNullOrWhiteSpace($DatasourceName)) { throw 'DatasourceName is required (or set projectWise.datasourceName in appsettings.json).' }
if ([string]::IsNullOrWhiteSpace($PwCredFilePath)) { throw 'PwCredFilePath is required (or set projectWise.credentialPath in appsettings.json).' }

$pwFolderPath = ConvertFrom-PWUriToFolderPath -Value $FolderPath
if ([string]::IsNullOrWhiteSpace($pwFolderPath)) {
    Write-Output (_FailJson -Code 'FOLDER_PATH_INVALID' -Message 'FolderPath could not be parsed.' -Data @{ input = $FolderPath })
    exit 2
}

$credRes = Get-PWCredentialFromFile -CredentialPath $PwCredFilePath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$cred = [pscredential]$credRes.Data.credential

$connRes = Connect-PW -DatasourceName $DatasourceName -Credential $cred
if (-not $connRes.IsSuccess) { throw $connRes.Message }

try {
    $cmd = Get-Command -Name 'Get-PWFolderTreeDocumentStateCount' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        Write-Output (_FailJson -Code 'MISSING_CMDLET' -Message 'Get-PWFolderTreeDocumentStateCount is not available in this runspace (pwps_dab required).' -Data @{ requiredCmdlet = 'Get-PWFolderTreeDocumentStateCount' })
        exit 3
    }

    $folderObj = $null
    try {
        $folderObj = Get-PWFolders -FolderPath $pwFolderPath -JustOne -ErrorAction Stop
    } catch {
        $folderObj = $null
    }
    if (-not $folderObj) {
        # Try without leading Documents\ for environments that prefer relative paths.
        if ($pwFolderPath -match '^(?i)Documents\\(.+)$') {
            try { $folderObj = Get-PWFolders -FolderPath $Matches[1] -JustOne -ErrorAction Stop } catch { $folderObj = $null }
        }
    }
    if (-not $folderObj) {
        Write-Output (_FailJson -Code 'FOLDER_NOT_FOUND' -Message 'Folder not resolved via Get-PWFolders -JustOne.' -Data @{ folderPath = $pwFolderPath })
        exit 4
    }

    $raw = Get-PWFolderTreeDocumentStateCount -InputFolder $folderObj

    $out = [ordered]@{
        ok = $true
        collectedAtUtc = Get-QCTimestamp
        datasourceName = $DatasourceName
        input = [ordered]@{
            folderPath = $FolderPath
            folderPathCanonical = $pwFolderPath
        }
        stateCounts = $raw
    }

    if ($Pretty) { Write-Output ($out | ConvertTo-Json -Depth 9) }
    else { Write-Output ($out | ConvertTo-Json -Depth 9 -Compress) }
} finally {
    $null = Disconnect-PW
}

