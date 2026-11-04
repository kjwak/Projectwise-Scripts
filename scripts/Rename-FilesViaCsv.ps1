<#
.SYNOPSIS
Export ProjectWise documents in a folder to a two-column CSV, then rename them from an edited copy.

.DESCRIPTION
Export writes one row per document in the folder:
  FullPath  - pw:\\server:datasource\Documents\... path (do not edit)
  FileName  - current ProjectWise file name (edit this column)

Apply reads that CSV and renames each document in place with Rename-PWDocument -RenameFile,
so the document name and the file name both become the edited FileName.
Rows whose FileName still matches the document are left unchanged.

Renames that take a name still held by another row in the same CSV run in dependency order:
the document that currently owns the target name is renamed first. A swap or other cycle is
parked on a temporary name, then moved to its final name. A retry treats a missing source
whose target name is already present as already applied.

.PARAMETER ProjectWiseFolder
ProjectWise folder. Accepts pw:\\server:datasource\Documents\... or a datasource-relative
path such as Caltrans\Project\CADD\Sheets\S_Seg. Export only.

.PARAMETER OutputCsv
CSV path written by Export.

.PARAMETER Recurse
Include documents in subfolders. Export only.

.PARAMETER Filter
File-name wildcard applied after the folder search. Default is all documents. Export only.

.PARAMETER InputCsv
Edited CSV to apply. Must contain FullPath and FileName columns.

.PARAMETER ResultCsv
Optional path for per-row apply results. Defaults to <InputCsv stem>.apply-results.csv.

.PARAMETER DatasourceName
ProjectWise datasource (server:name). Used when opening a new session.

.PARAMETER CredentialPath
Optional key=value credential file (username=/password=). Used only if no usable session exists
and SSO login fails. Secrets are never written to output files.

.PARAMETER KeepSession
Leave the ProjectWise session open if this script opened it.

.EXAMPLE
.\scripts\Rename-FilesViaCsv.ps1 -ProjectWiseFolder 'pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\S_Seg\' -OutputCsv 'D:\rename.csv'

.EXAMPLE
.\scripts\Rename-FilesViaCsv.ps1 -InputCsv 'D:\rename.csv' -WhatIf

.EXAMPLE
.\scripts\Rename-FilesViaCsv.ps1 -InputCsv 'D:\rename.csv'
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Export')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Export')]
    [string]$ProjectWiseFolder,

    [Parameter(Mandatory = $true, ParameterSetName = 'Export')]
    [string]$OutputCsv,

    [Parameter(Mandatory = $false, ParameterSetName = 'Export')]
    [switch]$Recurse,

    [Parameter(Mandatory = $false, ParameterSetName = 'Export')]
    [string]$Filter = '*',

    [Parameter(Mandatory = $true, ParameterSetName = 'Apply')]
    [string]$InputCsv,

    [Parameter(Mandatory = $false, ParameterSetName = 'Apply')]
    [string]$ResultCsv = '',

    [Parameter(Mandatory = $false)]
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',

    [Parameter(Mandatory = $false)]
    [string]$CredentialPath = '',

    [Parameter(Mandatory = $false)]
    [switch]$KeepSession
)

# ProjectWise native libraries (dmscli.dll) load in Windows PowerShell 5.1.
# PowerShell 7 can see the cmdlet and still fail to load that DLL.
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPs = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $winPs)) {
        throw 'ProjectWise cmdlets require Windows PowerShell 5.1, and powershell.exe was not found.'
    }
    $forward = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    foreach ($key in @($PSBoundParameters.Keys)) {
        $val = $PSBoundParameters[$key]
        if ($val -is [switch]) {
            if ($val.IsPresent) { $forward += "-$key" }
        }
        else {
            $forward += "-$key"
            $forward += [string]$val
        }
    }
    & $winPs @forward
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
# pwps owns Open-PWConnection. pwps_dab does not import it, and PowerShell 7 does not autoload it.
Import-Module pwps, pwps_dab -ErrorAction Stop
Set-StrictMode -Version Latest

function Get-PwSafeProperty {
    param(
        [object]$Object,
        [string]$Name
    )
    if ($null -eq $Object) { return $null }
    try {
        $prop = $Object.PSObject.Properties[$Name]
        if ($null -eq $prop) { return $null }
        return $prop.Value
    }
    catch {
        return $null
    }
}

function ConvertTo-PwFolderSearchPath {
    param([string]$RawPath)

    $p = ($RawPath -as [string]).Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($p)) {
        throw 'ProjectWise folder path is empty.'
    }

    if ($p -match '^pw:[\\/]{2}[^\\/]+[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }

    $p = $p.Trim().TrimStart('\').TrimEnd('\')
    if ($p -match '^(?i)Documents[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }
    return $p.TrimStart('\').TrimEnd('\')
}

function Split-PwDocumentUri {
    param([string]$RawPath)

    $p = ($RawPath -as [string]).Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($p)) {
        throw 'FullPath is empty.'
    }
    if ($p -match '^pw:[\\/]{2}[^\\/]+[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }
    $p = $p.Trim().TrimStart('\').TrimEnd('\')
    if ($p -match '^(?i)Documents[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }
    $p = $p.TrimStart('\').TrimEnd('\')
    $fileName = [System.IO.Path]::GetFileName($p)
    $folder = [System.IO.Path]::GetDirectoryName($p)
    if ([string]::IsNullOrWhiteSpace($fileName) -or [string]::IsNullOrWhiteSpace($folder)) {
        throw "FullPath is not a document path: $RawPath"
    }
    return [pscustomobject]@{
        FolderPath = ($folder -replace '/', '\').Trim('\')
        FileName   = $fileName
    }
}

function Format-PwDocumentUri {
    param(
        [string]$Datasource,
        [string]$FolderPath,
        [string]$FileName
    )
    $folder = ($FolderPath -replace '/', '\').Trim('\')
    if ($folder -notmatch '^(?i)Documents\\') {
        $folder = 'Documents\' + $folder
    }
    return ('pw:\\{0}\{1}\{2}' -f $Datasource.Trim(), $folder, $FileName)
}

function Test-SafeFileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return 'FileName is empty.'
    }
    if ($Name -match '[\\/]' -or $Name -eq '.' -or $Name -eq '..') {
        return 'FileName must be a file name only, not a path.'
    }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -contains $ch) {
            return "FileName contains an invalid character: '$ch'."
        }
    }
    return $null
}

function Test-PwActiveSession {
    param([string]$ExpectedDatasourceName)

    try {
        $null = Get-PWLoginStatus -ErrorAction Stop
    }
    catch {
        return $false
    }

    try {
        $ds = Get-PWDatasource -Active -ErrorAction Stop
        $full = [string](Get-PwSafeProperty -Object $ds -Name 'FullName')
        $internal = [string](Get-PwSafeProperty -Object $ds -Name 'InternalName')
        if ([string]::IsNullOrWhiteSpace($full) -and [string]::IsNullOrWhiteSpace($internal)) {
            return $true
        }
        if ($full -eq $ExpectedDatasourceName) { return $true }
        if ($internal -and $ExpectedDatasourceName -like ('*:' + $internal)) { return $true }
        throw "Active ProjectWise datasource '$full' does not match expected '$ExpectedDatasourceName'. The existing session was left unchanged."
    }
    catch {
        if ($_.Exception.Message -like 'Active ProjectWise datasource*') { throw }
        return $true
    }
}

function Get-PwCredentialFromKeyFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    $lines = Get-Content -LiteralPath $Path
    $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
    $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
    if (-not $uLine -or -not $pLine) {
        throw "Invalid credential file format: $Path"
    }
    $userName = ($uLine -split '=', 2)[1].Trim()
    $plain = ($pLine -split '=', 2)[1].Trim()
    if ([string]::IsNullOrWhiteSpace($userName) -or [string]::IsNullOrWhiteSpace($plain)) {
        throw "Credential file missing username/password: $Path"
    }
    $sec = ConvertTo-SecureString $plain -AsPlainText -Force
    return [pscredential]::new($userName, $sec)
}

function Get-PwDocumentFileName {
    param([object]$Document)
    foreach ($n in @('FileName', 'Name')) {
        $v = [string](Get-PwSafeProperty -Object $Document -Name $n)
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
    }
    return ''
}

function Resolve-PwFolder {
    param([string]$RawPath)

    $searchPath = ConvertTo-PwFolderSearchPath -RawPath $RawPath
    $folderErrors = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @($searchPath, ('Documents\' + $searchPath))) {
        try {
            $hit = Get-PWFolders -FolderPath $candidate -JustOne -ErrorAction Stop
            if ($hit) {
                $folder = @($hit)[0]
                $full = [string](Get-PwSafeProperty -Object $folder -Name 'FullPath')
                if ([string]::IsNullOrWhiteSpace($full)) {
                    try { $full = [string]$folder.GetFullPath() } catch { $full = $candidate }
                }
                return [pscustomobject]@{
                    SearchPath = $candidate
                    FullPath   = ($full -replace '/', '\').Trim('\')
                    Name       = [string](Get-PwSafeProperty -Object $folder -Name 'Name')
                }
            }
        }
        catch {
            [void]$folderErrors.Add("$candidate : $($_.Exception.Message)")
        }
    }
    throw ("Could not resolve ProjectWise folder '$RawPath'. Tried: " + ($folderErrors -join '; '))
}

function Get-PwDocumentsInFolder {
    param(
        [string]$FolderPath,
        [switch]$Recurse
    )

    $docs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
    if (-not $Recurse) {
        return $docs
    }

    $children = @(Get-PWFoldersImmediateChildren -FolderPath $FolderPath -ErrorAction Stop)
    foreach ($child in $children) {
        $childPath = [string](Get-PwSafeProperty -Object $child -Name 'FullPath')
        if ([string]::IsNullOrWhiteSpace($childPath)) {
            $childName = [string](Get-PwSafeProperty -Object $child -Name 'Name')
            $childPath = ($FolderPath.TrimEnd('\') + '\' + $childName)
        }
        $docs += @(Get-PwDocumentsInFolder -FolderPath $childPath -Recurse)
    }
    return $docs
}

function Get-PwRenameActions {
    param([System.Collections.Generic.List[object]]$Items)

    $pendingIds = New-Object 'System.Collections.Generic.HashSet[int]'
    $byId = @{}
    foreach ($item in $Items) {
        [void]$pendingIds.Add([int]$item.Id)
        $byId[[int]$item.Id] = $item
    }

    $reserved = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Items) {
        [void]$reserved.Add((($item.FolderPath + '\' + $item.CurrentName).ToLowerInvariant()))
        [void]$reserved.Add((($item.FolderPath + '\' + $item.NewName).ToLowerInvariant()))
    }

    $actions = New-Object System.Collections.Generic.List[object]
    $guard = 0
    $guardMax = ($Items.Count * 4) + 2

    while ($pendingIds.Count -gt 0) {
        $guard++
        if ($guard -gt $guardMax) {
            throw 'Could not order renames without looping. Check the CSV for a name cycle that cannot be parked.'
        }

        $ready = New-Object System.Collections.Generic.List[int]
        foreach ($id in @($pendingIds)) {
            $item = $byId[$id]
            $want = ($item.FolderPath + '\' + $item.NewName).ToLowerInvariant()
            $blocked = $false
            foreach ($otherId in @($pendingIds)) {
                if ($otherId -eq $id) { continue }
                $other = $byId[$otherId]
                if ($item.FolderPath -ne $other.FolderPath) { continue }
                $held = ($other.FolderPath + '\' + $other.CurrentName).ToLowerInvariant()
                if ($held -eq $want) {
                    $blocked = $true
                    break
                }
            }
            if (-not $blocked) { $ready.Add($id) }
        }

        if ($ready.Count -eq 0) {
            $id = @($pendingIds)[0]
            $item = $byId[$id]
            $ext = [System.IO.Path]::GetExtension($item.CurrentName)
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($item.CurrentName)
            $tempName = $null
            for ($n = 0; $n -lt 5 -and [string]::IsNullOrWhiteSpace($tempName); $n++) {
                $candidate = '{0}.__rn{1}{2}' -f $stem, [guid]::NewGuid().ToString('N').Substring(0, 8), $ext
                $candidateKey = ($item.FolderPath + '\' + $candidate).ToLowerInvariant()
                if (-not $reserved.Contains($candidateKey)) { $tempName = $candidate }
            }
            if ([string]::IsNullOrWhiteSpace($tempName)) {
                throw "Could not choose a temporary name for $($item.CurrentName)."
            }
            $actions.Add([pscustomobject]@{
                Id         = $item.Id
                FullPath   = $item.FullPath
                FolderPath = $item.FolderPath
                OldName    = $item.OldName
                FileName   = $item.NewName
                SourceName = $item.CurrentName
                DestName   = $tempName
                Phase      = 'Temp'
            })
            $oldHeld = ($item.FolderPath + '\' + $item.CurrentName).ToLowerInvariant()
            [void]$reserved.Remove($oldHeld)
            $item.CurrentName = $tempName
            [void]$reserved.Add((($item.FolderPath + '\' + $tempName).ToLowerInvariant()))
            continue
        }

        foreach ($id in @($ready | Sort-Object { $byId[$_].FolderPath + '\' + $byId[$_].NewName } -Descending)) {
            $item = $byId[$id]
            $actions.Add([pscustomobject]@{
                Id         = $item.Id
                FullPath   = $item.FullPath
                FolderPath = $item.FolderPath
                OldName    = $item.OldName
                FileName   = $item.NewName
                SourceName = $item.CurrentName
                DestName   = $item.NewName
                Phase      = 'Final'
            })
            [void]$pendingIds.Remove($id)
            $item.CurrentName = $item.NewName
        }
    }

    return $actions
}

function Find-PwDocumentByFileName {
    param(
        [string]$FolderPath,
        [string]$FileName
    )

    $byFile = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -FileName $FileName -PopulatePath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    if ($byFile.Count -eq 1) { return $byFile[0] }
    if ($byFile.Count -gt 1) {
        throw "More than one document in '$FolderPath' has file name '$FileName'."
    }

    $byDoc = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $FileName -PopulatePath -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    if ($byDoc.Count -eq 1) { return $byDoc[0] }
    if ($byDoc.Count -gt 1) {
        throw "More than one document in '$FolderPath' has document name '$FileName'."
    }
    return $null
}

$openedByScript = $false
try {
    $connected = Test-PwActiveSession -ExpectedDatasourceName $DatasourceName
    if (-not $connected) {
        $ssoFailed = $null
        try {
            Write-Host "No active ProjectWise session. Trying SSO for $DatasourceName ..."
            Open-PWConnection -DatasourceName $DatasourceName -SSO -WarningAction SilentlyContinue | Out-Null
            $connected = $true
            $openedByScript = $true
            Write-Host 'Connected via SSO.'
        }
        catch {
            $ssoFailed = $_.Exception.Message
        }

        if (-not $connected) {
            $credFile = $CredentialPath
            if ([string]::IsNullOrWhiteSpace($credFile) -and $env:PW_CRED_PATH) {
                $credFile = [string]$env:PW_CRED_PATH
            }
            if ([string]::IsNullOrWhiteSpace($credFile) -and (Test-Path -LiteralPath 'C:\PW_QC_LOCAL\pw_cred.txt')) {
                $credFile = 'C:\PW_QC_LOCAL\pw_cred.txt'
            }
            $cred = Get-PwCredentialFromKeyFile -Path $credFile
            if ($null -ne $cred) {
                Write-Host "SSO failed. Connecting with credential file (username not logged)."
                Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
                $connected = $true
                $openedByScript = $true
            }
        }

        if (-not $connected) {
            try {
                Write-Host 'Trying BentleyIMS ...'
                Open-PWConnection -DatasourceName $DatasourceName -BentleyIMS -WarningAction SilentlyContinue | Out-Null
                $connected = $true
                $openedByScript = $true
            }
            catch {
                throw ("Unable to open ProjectWise session. SSO: $ssoFailed; BentleyIMS: $($_.Exception.Message)")
            }
        }
    }
    else {
        Write-Host 'Using existing ProjectWise session.'
    }

    $activeDs = $null
    try { $activeDs = Get-PWDatasource -Active } catch { }
    $activeDsName = [string](Get-PwSafeProperty -Object $activeDs -Name 'FullName')
    if ([string]::IsNullOrWhiteSpace($activeDsName)) { $activeDsName = $DatasourceName }
    Write-Host "Datasource: $activeDsName"

    if ($PSCmdlet.ParameterSetName -eq 'Export') {
        $folder = Resolve-PwFolder -RawPath $ProjectWiseFolder
        Write-Host ("Resolved folder: {0}" -f $folder.FullPath)

        $docs = @(Get-PwDocumentsInFolder -FolderPath $folder.SearchPath -Recurse:$Recurse)
        $rows = @(
            $docs | ForEach-Object {
                $fileName = Get-PwDocumentFileName -Document $_
                if ([string]::IsNullOrWhiteSpace($fileName)) { return }
                if ($fileName -notlike $Filter) { return }
                $docFolder = [string](Get-PwSafeProperty -Object $_ -Name 'FolderPath')
                if ([string]::IsNullOrWhiteSpace($docFolder)) { $docFolder = $folder.FullPath }
                [pscustomobject]@{
                    FullPath = Format-PwDocumentUri -Datasource $activeDsName -FolderPath $docFolder -FileName $fileName
                    FileName = $fileName
                }
            } | Sort-Object FullPath
        )

        $outDir = Split-Path -Parent $OutputCsv
        if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        }
        $rows | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
        Write-Host ("Exported {0} document(s) to {1}" -f $rows.Count, $OutputCsv)
        Write-Host 'Edit the FileName column, then re-run with -InputCsv to apply renames.'
        return
    }

    $csvPath = (Resolve-Path -LiteralPath $InputCsv).Path
    $rows = @(Import-Csv -LiteralPath $csvPath)
    if ($rows.Count -eq 0) {
        throw "CSV has no data rows: $csvPath"
    }
    $headers = @($rows[0].PSObject.Properties.Name)
    foreach ($required in @('FullPath', 'FileName')) {
        if ($headers -notcontains $required) {
            throw "CSV is missing required column '$required'."
        }
    }

    if ([string]::IsNullOrWhiteSpace($ResultCsv)) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
        $ResultCsv = Join-Path (Split-Path -Parent $csvPath) ($stem + '.apply-results.csv')
    }

    $results = New-Object System.Collections.Generic.List[object]
    $work = New-Object System.Collections.Generic.List[object]
    $plannedTargets = @{}
    $keptNames = @{}
    $nextId = 0

    foreach ($row in $rows) {
        $fullPath = ([string]$row.FullPath).Trim()
        $newName = ([string]$row.FileName).Trim()
        $oldName = ''
        try {
            $nameError = Test-SafeFileName -Name $newName
            if ($nameError) { throw $nameError }
            $parsed = Split-PwDocumentUri -RawPath $fullPath
            $oldName = $parsed.FileName
            $targetKey = ($parsed.FolderPath + '\' + $newName).ToLowerInvariant()
            if ($plannedTargets.ContainsKey($targetKey)) {
                throw "Another row in this CSV already targets $targetKey."
            }
            $plannedTargets[$targetKey] = $fullPath
            if ($oldName -ceq $newName) {
                $keptNames[$targetKey] = $true
                $results.Add([pscustomobject]@{
                    Status   = 'Unchanged'
                    FullPath = $fullPath
                    OldName  = $oldName
                    FileName = $newName
                    Message  = 'FileName matches the current name.'
                }) | Out-Null
                continue
            }
            if ($keptNames.ContainsKey($targetKey)) {
                throw "Target name is kept by another row that is not being renamed: $newName"
            }
            $nextId++
            $work.Add([pscustomobject]@{
                Id          = $nextId
                FullPath    = $fullPath
                FolderPath  = $parsed.FolderPath
                OldName     = $oldName
                NewName     = $newName
                CurrentName = $oldName
            }) | Out-Null
        }
        catch {
            $results.Add([pscustomobject]@{
                Status   = 'Error'
                FullPath = $fullPath
                OldName  = $oldName
                FileName = $newName
                Message  = $_.Exception.Message
            }) | Out-Null
        }
    }

    $actions = @(Get-PwRenameActions -Items $work)
    $released = @{}

    foreach ($action in $actions) {
        $status = 'Skipped'
        $message = ''
        try {
            $sourceKey = ($action.FolderPath + '\' + $action.SourceName).ToLowerInvariant()
            $destKey = ($action.FolderPath + '\' + $action.DestName).ToLowerInvariant()
            $doc = Find-PwDocumentByFileName -FolderPath $action.FolderPath -FileName $action.SourceName
            if ($null -eq $doc) {
                if ($action.Phase -eq 'Final' -and $action.SourceName -ceq $action.OldName) {
                    $already = Find-PwDocumentByFileName -FolderPath $action.FolderPath -FileName $action.DestName
                    if ($null -ne $already) {
                        $status = 'AlreadyApplied'
                        $message = 'Source name is already gone and the target name is present.'
                        $released[$sourceKey] = $true
                    }
                    else {
                        $status = 'Error'
                        $message = "Document was not found: $($action.FullPath)"
                    }
                }
                else {
                    $status = 'Error'
                    $message = "Document '$($action.SourceName)' was not found in $($action.FolderPath)."
                }
            }
            else {
                $existing = Find-PwDocumentByFileName -FolderPath $action.FolderPath -FileName $action.DestName
                $sameDoc = $false
                if ($null -ne $existing) {
                    $existingId = [string](Get-PwSafeProperty -Object $existing -Name 'DocumentID')
                    $docId = [string](Get-PwSafeProperty -Object $doc -Name 'DocumentID')
                    $sameDoc = ($existingId -eq $docId -and -not [string]::IsNullOrWhiteSpace($docId))
                }
                $virtuallyFree = $released.ContainsKey($destKey)
                if ($null -ne $existing -and -not $sameDoc -and -not $virtuallyFree) {
                    $status = 'Error'
                    $message = "Target already exists in $($action.FolderPath): $($action.DestName)"
                }
                else {
                    $verb = if ($action.Phase -eq 'Temp') { "Park as $($action.DestName)" } else { "Rename ProjectWise document to $($action.DestName)" }
                    if ($PSCmdlet.ShouldProcess($action.FullPath, $verb)) {
                        Rename-PWDocument -InputDocument $doc -DocumentNewName $action.DestName -RenameFile -ErrorAction Stop | Out-Null
                        $status = if ($action.Phase -eq 'Temp') { 'Parked' } else { 'Renamed' }
                        $message = Format-PwDocumentUri -Datasource $activeDsName -FolderPath $action.FolderPath -FileName $action.DestName
                    }
                    else {
                        $status = 'WhatIf'
                        $message = Format-PwDocumentUri -Datasource $activeDsName -FolderPath $action.FolderPath -FileName $action.DestName
                    }
                    $released[$sourceKey] = $true
                    $released.Remove($destKey)
                }
            }
        }
        catch {
            $status = 'Error'
            $message = $_.Exception.Message
        }

        if ($action.Phase -eq 'Temp' -and $status -eq 'Error') {
            $results.Add([pscustomobject]@{
                Status   = 'Error'
                FullPath = $action.FullPath
                OldName  = $action.OldName
                FileName = $action.FileName
                Message  = "Temporary rename failed: $message"
            }) | Out-Null
            continue
        }
        if ($action.Phase -eq 'Temp' -and $status -ne 'Error') {
            continue
        }

        $results.Add([pscustomobject]@{
            Status   = $status
            FullPath = $action.FullPath
            OldName  = $action.OldName
            FileName = $action.FileName
            Message  = $message
        }) | Out-Null
    }

    $results | Export-Csv -LiteralPath $ResultCsv -NoTypeInformation -Encoding UTF8
    $counts = $results | Group-Object Status | Sort-Object Name
    Write-Host 'Apply summary:'
    foreach ($group in $counts) {
        Write-Host ("  {0}: {1}" -f $group.Name, $group.Count)
    }
    Write-Host "Results: $ResultCsv"
    $results
}
finally {
    if ($openedByScript -and -not $KeepSession) {
        try { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}
