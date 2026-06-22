<#
.SYNOPSIS
Collects read-only ProjectWise metrics for key project subfolders.

.DESCRIPTION
Targets these subfolders under a ProjectWise project root:
  - CADD
  - Proj_Management
  - Design_Engineering
  - Submittals

Collects fast rollups when available (pwps_dab folder-tree cmdlets), plus minimal identity info.
Outputs JSON to stdout (or to a file via -OutPath).

This script is intentionally read-only.
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

    # Example:
    # pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\AZDOT 2024\AZFWY...\   (trailing slash ok)
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [Parameter(Mandatory = $false)]
    [string]$OutPath
)

$ErrorActionPreference = 'Stop'

$scriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptPath)) { throw 'Cannot resolve script path via $MyInvocation.MyCommand.Path.' }
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptPath))

if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}
Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Config.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force

function ConvertFrom-PWUriToFolderPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)

    $s = ($Value -as [string]).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s = $s.Trim().Trim('"').Trim("'").TrimEnd('\')

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

function Invoke-PWFolderMetricCmdlet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CmdletName,
        [Parameter(Mandatory)][hashtable]$Arguments
    )

    $cmd = Get-Command -Name $CmdletName -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return @{ ok = $false; kind = 'missing_cmdlet'; cmdlet = $CmdletName }
    }

    try {
        $result = & $CmdletName @Arguments
        return @{ ok = $true; kind = 'ok'; cmdlet = $CmdletName; data = $result }
    } catch {
        return @{ ok = $false; kind = 'error'; cmdlet = $CmdletName; errorMessage = $_.Exception.Message }
    }
}

function Get-PWSubfolderMetrics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$Label
    )

    $folderPathCanon = $FolderPath.TrimEnd('\')

    $folderObj = $null
    $tryPaths = @($folderPathCanon)
    if ($folderPathCanon -match '^(?i)Documents\\(.+)$') {
        $tryPaths += $Matches[1]
    }
    foreach ($p in @($tryPaths | Select-Object -Unique)) {
        try {
            $folderObj = Get-PWFolders -FolderPath $p -JustOne -ErrorAction Stop
            if ($folderObj) { break }
        } catch { }
    }

    # Prefer fast rollups when present. Folder-tree cmdlets take -InputFolder (object), not -FolderPath (string).
    $maxUpdate = if ($folderObj) {
        Invoke-PWFolderMetricCmdlet -CmdletName 'Get-PWFolderTreeMaxFileUpdateDate' -Arguments @{ InputFolder = $folderObj }
    } else {
        @{ ok = $false; kind = 'missing_folder'; cmdlet = 'Get-PWFolderTreeMaxFileUpdateDate'; errorMessage = 'Folder not resolved (Get-PWFolders -JustOne returned nothing).' }
    }

    $treeSize = if ($folderObj) {
        Invoke-PWFolderMetricCmdlet -CmdletName 'Get-PWFolderTreeDocumentSize' -Arguments @{ InputFolder = $folderObj }
    } else {
        @{ ok = $false; kind = 'missing_folder'; cmdlet = 'Get-PWFolderTreeDocumentSize'; errorMessage = 'Folder not resolved (Get-PWFolders -JustOne returned nothing).' }
    }
    $sizeByPath = Invoke-PWFolderMetricCmdlet -CmdletName 'Get-PWFolderDocumentSizeByPath' -Arguments @{ FolderPath = $folderPathCanon }
    if (-not [bool]$sizeByPath.ok) {
        $pwUri = ('pw:\\' + $DatasourceName + '\' + $folderPathCanon.TrimStart('\')).TrimEnd('\')
        $sizeByPath = Invoke-PWFolderMetricCmdlet -CmdletName 'Get-PWFolderDocumentSizeByPath' -Arguments @{ FolderPath = $pwUri }
        if (-not [bool]$sizeByPath.ok -and $folderPathCanon -match '^(?i)Documents\\(.+)$') {
            $sizeByPath = Invoke-PWFolderMetricCmdlet -CmdletName 'Get-PWFolderDocumentSizeByPath' -Arguments @{ FolderPath = $Matches[1] }
        }
    }
    $stateCounts = if ($folderObj) {
        Invoke-PWFolderMetricCmdlet -CmdletName 'Get-PWFolderTreeDocumentStateCount' -Arguments @{ InputFolder = $folderObj }
    } else {
        @{ ok = $false; kind = 'missing_folder'; cmdlet = 'Get-PWFolderTreeDocumentStateCount'; errorMessage = 'Folder not resolved (Get-PWFolders -JustOne returned nothing).' }
    }

    return [ordered]@{
        label = $Label
        folderPath = $folderPathCanon
        metrics = [ordered]@{
            maxFileUpdateDate = $maxUpdate
            folderTreeDocumentSize = $treeSize
            documentSizeByPath = $sizeByPath
            documentStateCounts = $stateCounts
        }
    }
}

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

$projFolderPath = ConvertFrom-PWUriToFolderPath -Value $ProjectRoot
if ([string]::IsNullOrWhiteSpace($projFolderPath)) { throw 'ProjectRoot could not be parsed into a ProjectWise folder path.' }

$credRes = Get-PWCredentialFromFile -CredentialPath $PwCredFilePath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$cred = $credRes.Data.credential

$connRes = Connect-PW -DatasourceName $DatasourceName -Credential $cred
if (-not $connRes.IsSuccess) { throw $connRes.Message }

try {
    $targets = @(
        @{ label = 'CADD'; rel = 'CADD' },
        @{ label = 'Proj_Management'; rel = 'Proj_Management' },
        @{ label = 'Design_Engineering'; rel = 'Design_Engineering' },
        @{ label = 'Submittals'; rel = 'Submittals' }
    )

    $out = [ordered]@{
        collectedAtUtc = Get-QCTimestamp
        datasourceName = $DatasourceName
        projectRoot = $projFolderPath
        subfolders = @()
    }

    foreach ($t in $targets) {
        $path = ($projFolderPath.TrimEnd('\') + '\' + [string]$t.rel).TrimEnd('\')
        $out.subfolders += (Get-PWSubfolderMetrics -FolderPath $path -Label ([string]$t.label))
    }

    $json = ($out | ConvertTo-Json -Depth 8)
    if ([string]::IsNullOrWhiteSpace($OutPath)) {
        $json
    } else {
        $dir = Split-Path -Parent $OutPath
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Set-Content -LiteralPath $OutPath -Value $json -Encoding UTF8
        Write-Host ("Wrote: {0}" -f $OutPath) -ForegroundColor Green
    }
} finally {
    $null = Disconnect-PW
}

