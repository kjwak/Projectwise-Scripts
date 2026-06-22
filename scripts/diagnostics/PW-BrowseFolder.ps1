<#
.SYNOPSIS
Connects to ProjectWise and prints child folders/docs for a given folder path.

.DESCRIPTION
Read-only helper to validate folder paths in the current datasource. The ProjectWise
connection and folder-view compatibility logic lives in modules\PW.Connection.psm1.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',

    [Parameter(Mandatory = $false)]
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',

    [Parameter(Mandatory = $true)]
    [string]$FolderPath,

    [Parameter(Mandatory = $false)]
    [int]$Max = 40
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force

Show-PWFolderBrowser -DatasourceName $DatasourceName -CredentialPath $CredentialPath -FolderPath $FolderPath -Max $Max
