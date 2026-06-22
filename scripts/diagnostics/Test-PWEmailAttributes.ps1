<#
.SYNOPSIS
Read-only probe: verify EM_Designer_Email / EM_Reveiewer_Email extraction from a ProjectWise folder.
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
    [int]$SampleSize = 8,

    [Parameter(Mandatory = $false)]
    [switch]$VerboseDump
)

$ErrorActionPreference = 'Stop'
$AttrNames = @('EM_Designer_Email', 'EM_Reveiewer_Email', 'EM_Reviewer_Email')

function ConvertFrom-PWUri {
    param([string]$Value)
    $s = $Value.Trim().TrimEnd('\')
    if ($s -match '^(?i)pw:\\\\[^\\]+\\(.+)$') { $s = $Matches[1] }
    if ($s -notmatch '^(?i)Documents\\') { $s = 'Documents\' + $s.TrimStart('\') }
    return $s.TrimEnd('\')
}

function Get-PWCredFromKeyValueFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Credential file not found: $Path" }
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
    $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
    if (-not $uLine -or -not $pLine) { throw "Bad credential file format: $Path" }
    $user = ($uLine -split '=', 2)[1].Trim()
    $pass = ($pLine -split '=', 2)[1].Trim()
    return [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))
}

function Get-EAttrNameValue {
    param([object]$Attr)
    $n = $null
    $v = $null
    if ($Attr.PSObject.Properties['Name']) { $n = [string]$Attr.Name }
    elseif ($Attr.PSObject.Properties['ColumnName']) { $n = [string]$Attr.ColumnName }
    if ($Attr.PSObject.Properties['Value']) { $v = [string]$Attr.Value }
    elseif ($Attr.PSObject.Properties['AttributeValue']) { $v = [string]$Attr.AttributeValue }
    return @{ name = $n; value = $v }
}

Import-Module pwps -Force
Import-Module pwps_dab -Force

$cred = Get-PWCredFromKeyValueFile -Path $CredentialPath
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null

try {
    $resolved = ConvertFrom-PWUri -Value $FolderPath
    $tryPaths = @($resolved)
    if ($resolved -match '^(?i)Documents\\(.+)$') { $tryPaths += $Matches[1] }

    $folderPath = $null
    foreach ($p in ($tryPaths | Select-Object -Unique)) {
        try {
            $f = Get-PWFolders -FolderPath $p -JustOne -ErrorAction Stop
            if ($f) { $folderPath = $p; break }
        } catch { }
    }
    if (-not $folderPath) { throw "Folder not found. Tried: $($tryPaths -join ', ')" }

    $out = [ordered]@{
        ok = $true
        folderPath = $folderPath
        workflowStateCounts = @()
        sample = @()
        summary = @{}
    }

    $folderObj = Get-PWFolders -FolderPath $folderPath -JustOne -ErrorAction Stop
    if (Get-Command Get-PWFolderTreeDocumentStateCount -ErrorAction SilentlyContinue) {
        $out.workflowStateCounts = @(Get-PWFolderTreeDocumentStateCount -InputFolder $folderObj | ForEach-Object {
            [ordered]@{
                workflowName = $_.WorkflowName
                stateName = $_.StateName
                count = $_.NumberOfDocumentsInState
                contained = $_.NumberOfContainedDocuments
            }
        })
    }

    $docs = @(Get-PWDocumentsBySearch -FolderPath $folderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
    $out.summary['documentCount'] = $docs.Count
    $out.summary['pdfCount'] = @($docs | Where-Object { $_.Name -match '\.pdf$' }).Count

    if ($VerboseDump) {
        $inspect = @($docs | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First 1)
        if ($inspect.Count -eq 0) { $inspect = @($docs | Select-Object -First 1) }
        if ($inspect.Count -gt 0) {
            $d0 = $inspect[0]
            $ea0 = @(Get-PWDocumentEAttributes -DocumentID $d0.DocumentID -ProjectID $d0.ProjectID -ErrorAction Stop)
            $out.verboseDump = [ordered]@{
                document = $d0.Name
                eattrCount = $ea0.Count
                eattrObjects = @($ea0 | ForEach-Object {
                    $props = @{}
                    foreach ($p in $_.PSObject.Properties) { $props[$p.Name] = $p.Value }
                    $props
                })
            }
            if (Get-Command Get-PWEnvironmentColumns -ErrorAction SilentlyContinue) {
                $cols = @(Get-PWEnvironmentColumns -ErrorAction SilentlyContinue)
                $out.verboseDump['environmentColumnsMatching'] = @($cols | Where-Object {
                    $n = if ($_.Name) { [string]$_.Name } else { [string]$_.ColumnName }
                    $n -match 'EM_|Designer|Reviewer|Reveiew'
                } | Select-Object -First 30 | ForEach-Object {
                    $props = @{}
                    foreach ($p in $_.PSObject.Properties) { $props[$p.Name] = $p.Value }
                    $props
                })
            }
        }
    }

    $prefer = @($docs | Where-Object { $_.Name -match '\.pdf$' })
    if ($prefer.Count -gt 0) {
        $sample = @($prefer | Select-Object -First $SampleSize)
        $out.summary['sampleSource'] = 'pdf'
    } else {
        $sample = @($docs | Select-Object -First $SampleSize)
        $out.summary['sampleSource'] = 'all'
    }
    $designerPop = 0
    $reviewerTypoPop = 0
    $reviewerPop = 0

    foreach ($doc in $sample) {
        $row = [ordered]@{
            name = $doc.Name
            documentId = $doc.DocumentID
            workflow = if ($doc.PSObject.Properties['WorkflowName']) { $doc.WorkflowName } else { $null }
            state = if ($doc.PSObject.Properties['StateName']) { $doc.StateName } else { $null }
            attributes = [ordered]@{}
            methods = [ordered]@{}
        }

        try {
            $eattrs = @(Get-PWDocumentEAttributes -DocumentID $doc.DocumentID -ProjectID $doc.ProjectID -ErrorAction Stop)
            $row.methods['Get-PWDocumentEAttributes'] = "ok;count=$($eattrs.Count)"
            foreach ($a in $eattrs) {
                $nv = Get-EAttrNameValue -Attr $a
                if ($nv.name -in $AttrNames) { $row.attributes[$nv.name] = $nv.value }
                # Hashtable/dictionary-style e-attr payloads
                if ($a -is [System.Collections.IDictionary]) {
                    foreach ($key in $AttrNames) {
                        if ($a.Contains($key) -and -not $row.attributes[$key]) { $row.attributes[$key] = [string]$a[$key] }
                    }
                }
                # Property-bag fallback: match EM_* on any property name
                foreach ($p in $a.PSObject.Properties) {
                    if ($p.Name -in $AttrNames) { $row.attributes[$p.Name] = [string]$p.Value }
                }
            }
        } catch {
            $row.methods['Get-PWDocumentEAttributes'] = $_.Exception.Message
        }

        if (Get-Command Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue) {
            $cols = @('Name', 'DocumentID') + $AttrNames
            try {
                $with = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $folderPath -JustThisFolder -DocumentName $doc.Name -ColumnsToReturn $cols -ErrorAction Stop | Select-Object -First 1
                if ($with) {
                    $row.methods['Get-PWDocumentsBySearchWithReturnColumns'] = 'ok'
                    foreach ($cn in $AttrNames) {
                        if ($with.PSObject.Properties[$cn] -and [string]::IsNullOrWhiteSpace([string]$row.attributes[$cn])) {
                            $row.attributes[$cn] = [string]$with.$cn
                        }
                    }
                } else {
                    $row.methods['Get-PWDocumentsBySearchWithReturnColumns'] = 'no row'
                }
            } catch {
                $row.methods['Get-PWDocumentsBySearchWithReturnColumns'] = $_.Exception.Message
            }
        }

        if ($row.attributes['EM_Designer_Email']) { $designerPop++ }
        if ($row.attributes['EM_Reveiewer_Email']) { $reviewerTypoPop++ }
        if ($row.attributes['EM_Reviewer_Email']) { $reviewerPop++ }

        $out.sample += $row
    }

    $out.summary['sampleSize'] = $sample.Count
    $out.summary['EM_Designer_Email_populated'] = $designerPop
    $out.summary['EM_Reveiewer_Email_populated'] = $reviewerTypoPop
    $out.summary['EM_Reviewer_Email_populated'] = $reviewerPop
    $out.summary['extractionVerified'] = ($designerPop -gt 0) -or ($reviewerTypoPop -gt 0) -or ($reviewerPop -gt 0)

    $out | ConvertTo-Json -Depth 8
} finally {
    try { Close-PWConnection | Out-Null } catch { }
}
