<#
.SYNOPSIS
Lists documents in a ProjectWise folder (read-only).

.DESCRIPTION
Connects, runs the module-backed ProjectWise document listing, and prints counts plus a small sample.
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
    [int]$Max = 25
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force

Show-PWFolderDocumentList -DatasourceName $DatasourceName -CredentialPath $CredentialPath -FolderPath $FolderPath -Max $Max
