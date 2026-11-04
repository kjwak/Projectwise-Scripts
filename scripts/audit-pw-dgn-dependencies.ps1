<#
.SYNOPSIS
Read-only ProjectWise inventory of DGN files in one folder and their direct reference counts.

.DESCRIPTION
Enumerates *.dgn documents directly in a ProjectWise folder (no recursion), records the last
file updater display name, and counts immediate DGN/MicroStation reference attachments using
Get-PWReferenceFiles -ReferenceLevel 0 (ProjectWise reference index; no export/checkout).

Does not check out, check in, export, or modify ProjectWise documents.

.PARAMETER ProjectWiseFolder
ProjectWise folder path. Accepts pw:\\server:datasource\Documents\... or a datasource-relative
path such as Caltrans\Project\CADD\Ref.

.PARAMETER OutputCsv
Path for the inventory CSV (DGN_Name,Reference_Count,Last_Modified_By).

.PARAMETER ErrorCsv
Path for per-document inspection errors. Written only if at least one error occurs.

.PARAMETER DatasourceName
ProjectWise datasource (server:name). Used when opening a new session.

.PARAMETER CredentialPath
Optional key=value credential file (username=/password=). Used only if no usable session exists
and SSO login fails. Secrets are never written to output files.

.PARAMETER KeepSession
Leave the ProjectWise session open if this script opened it.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectWiseFolder,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = (Join-Path (Get-Location).Path 'dgn_dependency_audit.csv'),

    [Parameter(Mandatory = $false)]
    [string]$ErrorCsv = (Join-Path (Get-Location).Path 'dgn_dependency_audit_errors.csv'),

    [Parameter(Mandatory = $false)]
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',

    [Parameter(Mandatory = $false)]
    [string]$CredentialPath = '',

    [Parameter(Mandatory = $false)]
    [switch]$KeepSession
)

$ErrorActionPreference = 'Stop'
Import-Module pwps_dab -ErrorAction Stop
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
        throw 'ProjectWiseFolder is empty.'
    }

    if ($p -match '^pw:[\\/]{2}[^\\/]+[\\/]Documents[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }
    elseif ($p -match '^pw:[\\/]{2}[^\\/]+[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }

    $p = $p.Trim().TrimStart('\').TrimEnd('\')
    if ($p -match '^(?i)Documents[\\/](?<rest>.+)$') {
        $p = $Matches['rest']
    }
    return $p.TrimStart('\').TrimEnd('\')
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

function Get-PwDocumentDisplayName {
    param([object]$Document)

    foreach ($n in @('Name', 'FileName', 'DocumentName', 'OldName')) {
        $v = [string](Get-PwSafeProperty -Object $Document -Name $n)
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
    }
    return ''
}

function Test-PwIsDgnName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    return ($Name -like '*.dgn')
}

function Get-PwLastModifiedByDisplayName {
    param([object]$Document)

    $full = $Document
    try {
        $gp = $Document.GetGeneralProperties()
        if ($null -ne $gp) { $full = $gp }
    }
    catch {
        # Fall back to the search object.
    }

    # File updater is the last modifier of document content (PW "File Updated By").
    foreach ($userProp in @('FileUpdater', 'DocumentUpdater')) {
        $userObj = Get-PwSafeProperty -Object $full -Name $userProp
        $desc = [string](Get-PwSafeProperty -Object $userObj -Name 'Description')
        if (-not [string]::IsNullOrWhiteSpace($desc)) { return $desc.Trim() }
        $uname = [string](Get-PwSafeProperty -Object $userObj -Name 'UserName')
        if ([string]::IsNullOrWhiteSpace($uname)) {
            $uname = [string](Get-PwSafeProperty -Object $userObj -Name 'Name')
        }
        if (-not [string]::IsNullOrWhiteSpace($uname)) { return $uname.Trim() }
    }

    foreach ($nameProp in @('FileUpdaterName', 'DocumentUpdaterName')) {
        $n = [string](Get-PwSafeProperty -Object $full -Name $nameProp)
        if (-not [string]::IsNullOrWhiteSpace($n)) { return $n.Trim() }
    }

    return $null
}

function Get-PwDirectReferenceCount {
    param([object]$Document)

    $prevWarn = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    try {
        $refs = @(Get-PWReferenceFiles -InputDocument $Document -ReferenceLevel 0 -ErrorAction Stop)
        if ($null -eq $refs) { return 0 }
        return @($refs).Count
    }
    finally {
        $WarningPreference = $prevWarn
    }
}

# --- Connect (read-only session) ---
$openedByScript = $false
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
            Write-Host "SSO failed ($ssoFailed). Connecting with credential file (username not logged)."
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
$hostName = [string](Get-PwSafeProperty -Object $activeDs -Name 'HostName')
$internalName = [string](Get-PwSafeProperty -Object $activeDs -Name 'InternalName')
Write-Host ("Datasource: $activeDsName  host=$hostName  internal=$internalName")

$searchPath = ConvertTo-PwFolderSearchPath -RawPath $ProjectWiseFolder
$folder = $null
$folderErrors = New-Object System.Collections.Generic.List[string]
foreach ($candidate in @($searchPath, ('Documents\' + $searchPath))) {
    try {
        $hit = Get-PWFolders -FolderPath $candidate -JustOne -ErrorAction Stop
        if ($hit) {
            $folder = @($hit)[0]
            $searchPath = $candidate
            break
        }
    }
    catch {
        [void]$folderErrors.Add("$candidate : $($_.Exception.Message)")
    }
}

if ($null -eq $folder) {
    throw ("Could not resolve ProjectWise folder '$ProjectWiseFolder'. Tried: " + ($folderErrors -join '; '))
}

$folderName = [string](Get-PwSafeProperty -Object $folder -Name 'Name')
$folderId = Get-PwSafeProperty -Object $folder -Name 'ProjectID'
$folderGuid = [string](Get-PwSafeProperty -Object $folder -Name 'ProjectGUIDString')
if ([string]::IsNullOrWhiteSpace($folderGuid)) {
    $folderGuid = [string](Get-PwSafeProperty -Object $folder -Name 'ProjectGUID')
}
$folderFull = [string](Get-PwSafeProperty -Object $folder -Name 'FullPath')
if ([string]::IsNullOrWhiteSpace($folderFull)) {
    try { $folderFull = [string]$folder.GetFullPath() } catch { $folderFull = $searchPath }
}

Write-Host ("Resolved folder: name=$folderName  path=$folderFull  id=$folderId  guid=$folderGuid")

$allDocs = @(Get-PWDocumentsBySearch -FolderPath $searchPath -JustThisFolder -PopulatePath -ErrorAction Stop)
$dgnDocs = @(
    $allDocs | Where-Object {
        (Test-PwIsDgnName -Name (Get-PwDocumentDisplayName -Document $_))
    }
)
$discoveredCount = @($dgnDocs).Count
Write-Host ("Documents in folder: $(@($allDocs).Count); DGN files: $discoveredCount")

$rows = New-Object System.Collections.ArrayList
$errors = New-Object System.Collections.ArrayList
$zeroRefs = 0
$positiveRefs = 0
$totalAttachments = 0
$refUnknown = 0
$userUnknown = 0

foreach ($doc in $dgnDocs) {
    $dgnName = Get-PwDocumentDisplayName -Document $doc
    if ([string]::IsNullOrWhiteSpace($dgnName)) {
        $dgnName = ('(unnamed-documentid-' + [string](Get-PwSafeProperty -Object $doc -Name 'DocumentID') + ')')
    }

    $refCountOut = ''
    $lastBy = ''

    try {
        $resolvedName = Get-PwLastModifiedByDisplayName -Document $doc
        if ([string]::IsNullOrWhiteSpace($resolvedName)) {
            $userUnknown++
            [void]$errors.Add([pscustomobject]@{
                    DGN_Name = $dgnName
                    Stage    = 'LastModifiedBy'
                    Error    = 'FileUpdater and DocumentUpdater were empty after GetGeneralProperties.'
                })
        }
        else {
            $lastBy = $resolvedName
        }
    }
    catch {
        $userUnknown++
        [void]$errors.Add([pscustomobject]@{
                DGN_Name = $dgnName
                Stage    = 'LastModifiedBy'
                Error    = $_.Exception.Message
            })
    }

    try {
        $n = [int](Get-PwDirectReferenceCount -Document $doc)
        $refCountOut = $n
        $totalAttachments += $n
        if ($n -eq 0) { $zeroRefs++ } else { $positiveRefs++ }
    }
    catch {
        $refUnknown++
        $refCountOut = ''
        [void]$errors.Add([pscustomobject]@{
                DGN_Name = $dgnName
                Stage    = 'ReferenceInspection'
                Error    = $_.Exception.Message
            })
    }

    [void]$rows.Add([pscustomobject]@{
            DGN_Name         = $dgnName
            Reference_Count  = $refCountOut
            Last_Modified_By = $lastBy
        })
}

$sorted = @($rows | Sort-Object -Property DGN_Name)
if (@($sorted).Count -ne $discoveredCount) {
    throw "Validation failed: discovered DGN count ($discoveredCount) != sorted row count ($(@($sorted).Count))."
}

$outputDir = Split-Path -Parent $OutputCsv
if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}
$sorted | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

$errorPathWritten = $null
if (@($errors).Count -gt 0) {
    $errDir = Split-Path -Parent $ErrorCsv
    if (-not [string]::IsNullOrWhiteSpace($errDir) -and -not (Test-Path -LiteralPath $errDir)) {
        New-Item -ItemType Directory -Path $errDir | Out-Null
    }
    @($errors | Sort-Object DGN_Name, Stage) | Export-Csv -LiteralPath $ErrorCsv -NoTypeInformation -Encoding UTF8
    $errorPathWritten = $ErrorCsv
}
elseif (Test-Path -LiteralPath $ErrorCsv) {
    Remove-Item -LiteralPath $ErrorCsv -Force
}

$csvRows = @(Import-Csv -LiteralPath $OutputCsv)
if (@($csvRows).Count -ne $discoveredCount) {
    throw "Validation failed: discovered DGN count ($discoveredCount) != CSV data rows ($(@($csvRows).Count))."
}

Write-Host ''
Write-Host '=== DGN dependency audit summary ==='
Write-Host ("Resolved folder: $folderFull")
Write-Host ("Folder search path: $searchPath")
Write-Host ("Datasource: $activeDsName")
Write-Host ("DGN files discovered: $discoveredCount")
Write-Host ("CSV data rows: $(@($csvRows).Count)")
Write-Host ("DGNs with 0 direct references: $zeroRefs")
Write-Host ("DGNs with 1+ direct references: $positiveRefs")
Write-Host ("Total direct reference attachments counted: $totalAttachments")
Write-Host ("Reference_Count unknown: $refUnknown")
Write-Host ("Last_Modified_By unknown: $userUnknown")
Write-Host ("Output: $OutputCsv")
if ($errorPathWritten) {
    Write-Host ("Errors: $errorPathWritten ($($errors.Count) row(s))")
}
else {
    Write-Host 'Errors: none'
}
Write-Host 'Reference cmdlet: Get-PWReferenceFiles -ReferenceLevel 0'
Write-Host 'Last modified property: FileUpdater.Description (fallback FileUpdaterName, then DocumentUpdater)'

if ($openedByScript -and -not $KeepSession) {
    try { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null } catch { }
}
