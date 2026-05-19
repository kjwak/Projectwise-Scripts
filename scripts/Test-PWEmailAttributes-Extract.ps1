<#
.SYNOPSIS
Read-only extraction of Caltrans EM_Designer_Email / EM_Reviewer_Email from a ProjectWise folder.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$FolderPath,

    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',
    [int]$SampleSize = 10,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$TargetAttrs = @('EM_Designer_Email', 'EM_Reviewer_Email')

function Get-DocumentAttributeMap {
    param([object]$DocRow)
    $map = @{}
    if (-not $DocRow -or -not $DocRow.Attributes) { return $map }
    foreach ($bag in @($DocRow.Attributes)) {
        if ($bag -is [System.Collections.IDictionary]) {
            foreach ($k in $bag.Keys) { $map[[string]$k] = [string]$bag[$k] }
        }
    }
    return $map
}

function ConvertFrom-PWUri {
    param([string]$Value)
    $s = $Value.Trim().TrimEnd('\')
    if ($s -match '^(?i)pw:\\\\[^\\]+\\(.+)$') { $s = $Matches[1] }
    if ($s -notmatch '^(?i)Documents\\') { $s = 'Documents\' + $s.TrimStart('\') }
    return $s.TrimEnd('\')
}

$lines = Get-Content -LiteralPath $CredentialPath
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred = [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))

Import-Module pwps, pwps_dab -Force
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null

try {
    $folderPath = ConvertFrom-PWUri -Value $FolderPath
    $tryPaths = @($folderPath)
    if ($folderPath -match '^(?i)Documents\\(.+)$') { $tryPaths += $Matches[1] }
    $resolved = $null
    foreach ($p in ($tryPaths | Select-Object -Unique)) {
        if (Get-PWFolders -FolderPath $p -JustOne -ErrorAction SilentlyContinue) { $resolved = $p; break }
    }
    if (-not $resolved) { throw "Folder not found: $folderPath" }

    $folder = Get-PWFolders -FolderPath $resolved -JustOne -ErrorAction Stop
    $stateCounts = @()
    if (Get-Command Get-PWFolderTreeDocumentStateCount -ErrorAction SilentlyContinue) {
        $stateCounts = @(Get-PWFolderTreeDocumentStateCount -InputFolder $folder | ForEach-Object {
            [ordered]@{ workflow = $_.WorkflowName; state = $_.StateName; count = $_.NumberOfDocumentsInState }
        })
    }

    $docs = @(Get-PWDocumentsBySearch -FolderPath $resolved -JustThisFolder -PopulatePath)
    $pdfs = @($docs | Where-Object { $_.Name -match '\.pdf$' })

    $designerPop = 0
    $reviewerPop = 0
    $samples = @()

    foreach ($d in $pdfs) {
        $row = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $resolved -JustThisFolder -DocumentName $d.Name -ColumnsToReturn $TargetAttrs -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $row) { continue }
        $attrs = Get-DocumentAttributeMap -DocRow $row
        $designer = if ($attrs.ContainsKey('EM_Designer_Email')) { $attrs['EM_Designer_Email'] } else { '' }
        $reviewer = if ($attrs.ContainsKey('EM_Reviewer_Email')) { $attrs['EM_Reviewer_Email'] } else { '' }
        if ($designer) { $designerPop++ }
        if ($reviewer) { $reviewerPop++ }
        if (($designer -or $reviewer) -and $samples.Count -lt $SampleSize) {
            $samples += [ordered]@{
                name = $d.Name
                workflowState = if ($row.WorkflowState) { [string]$row.WorkflowState } else { $null }
                EM_Designer_Email = $designer
                EM_Reviewer_Email = $reviewer
            }
        }
    }

    $out = [ordered]@{
        ok = $true
        folderPath = $resolved
        environment = [string]$folder.Environment
        workflow = [string]$folder.Workflow
        workflowStateCounts = $stateCounts
        summary = [ordered]@{
            pdfCount = $pdfs.Count
            EM_Designer_Email_populated = $designerPop
            EM_Reviewer_Email_populated = $reviewerPop
            extractionVerified = ($designerPop -gt 0) -or ($reviewerPop -gt 0)
            note = 'Read via Get-PWDocumentsBySearchWithReturnColumns; values are in returned .Attributes SortedList bags.'
        }
        samples = $samples
    }

    if ($Pretty) { $out | ConvertTo-Json -Depth 8 } else { $out | ConvertTo-Json -Depth 8 -Compress }
} finally {
    Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
