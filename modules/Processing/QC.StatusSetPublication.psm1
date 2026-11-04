# QC.StatusSetPublication.psm1
# Copies configured Status Set PDFs to a ProjectWise archive folder.
# Scheduled runs and Test-StatusSetPublication.ps1 both call Invoke-StatusSetPublication.
# Source PDFs are never updated or deleted.

Import-Module (Join-Path $PSScriptRoot '..\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\Core\Core.Paths.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.StatusSet.psm1') -Force

$script:SSPubPendingFileName = 'statusset-publication-pending.json'
$script:SSPubPendingMaxAgeHours = 36

function _SSPub-Sha256File([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $hash = $sha.ComputeHash($fs)
    } finally {
        if ($fs) { $fs.Dispose() }
        $sha.Dispose()
    }
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function _SSPub-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _SSPub-AsHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in @($Value.Keys)) { $h[[string]$k] = $Value[$k] }
        return $h
    }
    if ($Value -is [pscustomobject]) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _SSPub-AsList([object]$Value) {
    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value) }
    return @($Value)
}

function _SSPub-Log {
    param(
        [scriptblock]$LogCallback,
        [string]$Level,
        [string]$Code,
        [string]$Message,
        [hashtable]$Data
    )
    if ($LogCallback) {
        try { & $LogCallback $Level $Code $Message $Data } catch { }
        return
    }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        try {
            Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $Data
        } catch { }
    }
}

function Resolve-StatusSetPublicationTimeZone {
    <#
    .SYNOPSIS
    Resolves America/Phoenix to the Windows zone used on the QC hosts.
    #>
    [CmdletBinding()]
    param([string]$TimeZoneId = 'America/Phoenix')

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not (_SSPub-IsBlank $TimeZoneId)) { [void]$candidates.Add($TimeZoneId.Trim()) }
    $lower = ''
    if (-not (_SSPub-IsBlank $TimeZoneId)) { $lower = $TimeZoneId.Trim().ToLowerInvariant() }
    if ($lower -in @('america/phoenix', 'us/arizona', 'arizona')) {
        [void]$candidates.Add('US Mountain Standard Time')
    }
    if ($candidates -notcontains 'US Mountain Standard Time') {
        [void]$candidates.Add('US Mountain Standard Time')
    }
    foreach ($id in @($candidates)) {
        try {
            $tz = [TimeZoneInfo]::FindSystemTimeZoneById($id)
            return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_TZ' -Message 'Publication time zone resolved.' -Data @{
                timeZone = $tz
                requestedId = $TimeZoneId
                resolvedId = $tz.Id
            }
        } catch { }
    }
    return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_TZ' -Message 'Could not resolve the publication time zone.' -Data @{
        requestedId = $TimeZoneId
    }
}

function ConvertTo-StatusSetPublicationFolderKey {
    <#
    .SYNOPSIS
    Comparable folder key: documents\... lowercased, pw:\\ datasource prefix removed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)

    if (_SSPub-IsBlank $Path) { return '' }
    $normalized = $Path.Trim().Replace('/', '\')
    while ($normalized.Contains('\\')) { $normalized = $normalized.Replace('\\', '\') }
    if ($normalized -match '^(?i)pw:\\') {
        $idx = $normalized.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) { $normalized = $normalized.Substring($idx + 1) }
    }
    $normalized = $normalized.Trim('\')
    if ($normalized -notmatch '^(?i)documents\\') {
        $normalized = 'Documents\' + $normalized
    }
    return $normalized.ToLowerInvariant()
}

function _SSPub-DocumentsDisplayPath([string]$Path) {
    $key = ConvertTo-StatusSetPublicationFolderKey -Path $Path
    if (_SSPub-IsBlank $key) { return '' }
    # Restore a Documents\ prefix in the case ProjectWise searches already use.
    if ($key.StartsWith('documents\')) {
        return ('Documents\' + $key.Substring('documents\'.Length))
    }
    return $key
}

function _SSPub-FolderTryPaths([string]$Path) {
    # Status-set uploads resolve Caltrans\... . A Documents\ prefix and a pw:\\ path
    # do not resolve that folder. New-PWDocument on an unresolved path returns null
    # and does not throw, so the stripped path has to be tried first.
    $display = _SSPub-DocumentsDisplayPath -Path $Path
    $ordered = New-Object System.Collections.Generic.List[string]
    $stripped = ''
    if ($display -match '^(?i)Documents\\') {
        $stripped = $display.Substring('Documents\'.Length)
    }
    if (-not (_SSPub-IsBlank $stripped)) { [void]$ordered.Add($stripped) }
    if (-not (_SSPub-IsBlank $display) -and -not $ordered.Contains($display)) { [void]$ordered.Add($display) }
    $original = ($Path -as [string]).Trim().TrimEnd('\')
    if (-not (_SSPub-IsBlank $original) -and -not $ordered.Contains($original)) { [void]$ordered.Add($original) }
    return @($ordered)
}

function _SSPub-PwFolderExists([string]$Path) {
    $cmd = Get-Command -Name 'Get-PWFolders' -ErrorAction SilentlyContinue
    if (-not $cmd) { return $true }
    try {
        $folder = Get-PWFolders -FolderPath $Path -JustOne -ErrorAction SilentlyContinue
        return ($null -ne $folder)
    } catch {
        return $false
    }
}

function Get-StatusSetPublicationSettings {
    <#
    .SYNOPSIS
    Normalizes statusSetPublication. Missing configuration yields no projects and no work.
    #>
    [CmdletBinding()]
    param([hashtable]$Config)

    $empty = @{
        enabled = $false
        timeZoneId = 'America/Phoenix'
        dayOfWeek = 'Monday'
        time = '06:00'
        projects = @()
        statusSetPdfName = '_StatusSet.pdf'
    }
    if (-not $Config) {
        return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_SETTINGS' -Message 'No configuration. Publication does no work.' -Data $empty
    }
    $pdfName = '_StatusSet.pdf'
    $ss = _SSPub-AsHashtable $Config['statusSet']
    if ($ss -and $ss.ContainsKey('statusSetPdfName') -and -not (_SSPub-IsBlank $ss['statusSetPdfName'])) {
        $pdfName = [string]$ss['statusSetPdfName']
    }
    $empty.statusSetPdfName = $pdfName

    if (-not $Config.ContainsKey('statusSetPublication') -or -not $Config['statusSetPublication']) {
        return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_SETTINGS' -Message 'statusSetPublication is not configured. Publication does no work.' -Data $empty
    }
    $raw = _SSPub-AsHashtable $Config['statusSetPublication']
    if (-not $raw) {
        return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_CONFIG' -Message 'statusSetPublication must be an object.' -Data @{}
    }

    $enabled = $false
    if ($raw.ContainsKey('enabled') -and $null -ne $raw['enabled']) {
        try { $enabled = [bool]$raw['enabled'] } catch { $enabled = $false }
    }
    $tzId = 'America/Phoenix'
    if ($raw.ContainsKey('timeZoneId') -and -not (_SSPub-IsBlank $raw['timeZoneId'])) { $tzId = [string]$raw['timeZoneId'] }
    $day = 'Monday'
    $time = '06:00'
    $trigger = _SSPub-AsHashtable $raw['trigger']
    if ($trigger) {
        if ($trigger.ContainsKey('dayOfWeek') -and -not (_SSPub-IsBlank $trigger['dayOfWeek'])) { $day = [string]$trigger['dayOfWeek'] }
        if ($trigger.ContainsKey('time') -and -not (_SSPub-IsBlank $trigger['time'])) { $time = [string]$trigger['time'] }
    }

    $projects = New-Object System.Collections.Generic.List[hashtable]
    foreach ($item in @(_SSPub-AsList $raw['projects'])) {
        $p = _SSPub-AsHashtable $item
        if (-not $p) { continue }
        $key = ''
        if ($p.ContainsKey('projectKey')) { $key = [string]$p['projectKey'] }
        $pEnabled = $true
        if ($p.ContainsKey('enabled') -and $null -ne $p['enabled']) {
            try { $pEnabled = [bool]$p['enabled'] } catch { $pEnabled = $true }
        }
        $sets = New-Object System.Collections.Generic.List[hashtable]
        foreach ($setItem in @(_SSPub-AsList $p['sets'])) {
            $set = _SSPub-AsHashtable $setItem
            if (-not $set) { continue }
            $ident = _SSPub-AsHashtable $set['identification']
            $method = 'statusSetPdfName'
            $documentName = ''
            if ($ident) {
                if ($ident.ContainsKey('method') -and -not (_SSPub-IsBlank $ident['method'])) { $method = [string]$ident['method'] }
                if ($ident.ContainsKey('documentName') -and -not (_SSPub-IsBlank $ident['documentName'])) { $documentName = [string]$ident['documentName'] }
            }
            if ($method -eq 'statusSetPdfName' -and (_SSPub-IsBlank $documentName)) { $documentName = $pdfName }
            $sets.Add(@{
                id = if ($set.ContainsKey('id')) { [string]$set['id'] } else { '' }
                sourceFolder = if ($set.ContainsKey('sourceFolder')) { [string]$set['sourceFolder'] } else { '' }
                destinationFolder = if ($set.ContainsKey('destinationFolder')) { [string]$set['destinationFolder'] } else { '' }
                fileNamePattern = if ($set.ContainsKey('fileNamePattern')) { [string]$set['fileNamePattern'] } else { '' }
                identificationMethod = $method
                documentName = $documentName
            }) | Out-Null
        }
        $projects.Add(@{
            projectKey = $key.Trim()
            enabled = $pEnabled
            sets = @($sets)
        }) | Out-Null
    }

    return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_SETTINGS' -Message 'Publication settings loaded.' -Data @{
        enabled = $enabled
        timeZoneId = $tzId
        dayOfWeek = $day
        time = $time
        projects = @($projects)
        statusSetPdfName = $pdfName
    }
}

function _SSPub-FindProject([hashtable]$Settings, [string]$ProjectKey) {
    foreach ($p in @($Settings.projects)) {
        if ([string]$p.projectKey -eq $ProjectKey.Trim()) { return $p }
    }
    foreach ($p in @($Settings.projects)) {
        if ([string]$p.projectKey -ieq $ProjectKey.Trim()) { return $p }
    }
    return $null
}

function _SSPub-ValidateSet([hashtable]$Set) {
    $problems = New-Object System.Collections.Generic.List[string]
    if (_SSPub-IsBlank $Set.id) { [void]$problems.Add('id is required') }
    if (_SSPub-IsBlank $Set.sourceFolder) { [void]$problems.Add('sourceFolder is required') }
    if (_SSPub-IsBlank $Set.destinationFolder) { [void]$problems.Add('destinationFolder is required') }
    if (_SSPub-IsBlank $Set.fileNamePattern) { [void]$problems.Add('fileNamePattern is required') }
    elseif ($Set.fileNamePattern -notmatch '\{yyyyMMdd\}' -and $Set.fileNamePattern -notmatch '\{yyyy-MM-dd\}') {
        [void]$problems.Add('fileNamePattern must include {yyyyMMdd} or {yyyy-MM-dd}')
    }
    $method = [string]$Set.identificationMethod
    if ($method -notin @('statusSetPdfName', 'documentName')) {
        [void]$problems.Add("identification method '$method' is not supported")
    }
    if ($method -eq 'documentName' -and (_SSPub-IsBlank $Set.documentName)) {
        [void]$problems.Add('identification.documentName is required')
    }
    if ($method -eq 'statusSetPdfName' -and (_SSPub-IsBlank $Set.documentName)) {
        [void]$problems.Add('status set PDF name is empty')
    }
    return @($problems)
}

function Format-StatusSetPublicationFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Pattern,
        [Parameter(Mandatory)][datetime]$AsOfDate,
        [string]$SetId = ''
    )
    $day = $AsOfDate.Date
    $name = $Pattern
    $name = $name.Replace('{yyyyMMdd}', $day.ToString('yyyyMMdd'))
    $name = $name.Replace('{yyyy-MM-dd}', $day.ToString('yyyy-MM-dd'))
    $name = $name.Replace('{segmentId}', [string]$SetId)
    return $name
}

function Test-StatusSetPublicationSlot {
    <#
    .SYNOPSIS
    True only for the configured weekday and time on a full_scan_schedule slot key.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Settings,
        [string]$SlotKey = ''
    )

    $no = @{
        matches = $false
        publicationDate = $null
        cycleStartUtc = $null
        reason = 'not_a_schedule_slot'
        resolvedTimeZoneId = ''
    }
    if (_SSPub-IsBlank $SlotKey) { $no.reason = 'no_slot'; return $no }
    if ($SlotKey -like 'ops_request|*') { $no.reason = 'ops_request'; return $no }
    if ($SlotKey -notmatch '^full_scan_schedule\|(\d{4}-\d{2}-\d{2})\|(\d{2}:\d{2})$') {
        $no.reason = 'unrecognized_slot'
        return $no
    }
    $dayText = $Matches[1]
    $hhmm = $Matches[2]
    $wantTime = ([string]$Settings.time).Trim()
    if ($wantTime.Length -eq 4 -and $wantTime -match '^\d{2}:\d{2}$' -eq $false) { }
    if ($hhmm -ne $wantTime) {
        $no.reason = 'time_mismatch'
        return $no
    }
    $tzRes = Resolve-StatusSetPublicationTimeZone -TimeZoneId ([string]$Settings.timeZoneId)
    if (-not $tzRes.IsSuccess) {
        $no.reason = 'timezone'
        return $no
    }
    $tz = $tzRes.Data.timeZone
    $no.resolvedTimeZoneId = [string]$tz.Id
    try {
        $localDate = [datetime]::ParseExact($dayText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        $no.reason = 'bad_date'
        return $no
    }
    $wantDay = [string]$Settings.dayOfWeek
    if ($localDate.DayOfWeek.ToString() -ne $wantDay) {
        $no.reason = 'weekday_mismatch'
        $no.publicationDate = $localDate.Date
        return $no
    }
    $localSlot = [datetime]::ParseExact(('{0} {1}' -f $dayText, $hhmm), 'yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
    $unspecified = [datetime]::SpecifyKind($localSlot, [DateTimeKind]::Unspecified)
    $cycleUtc = [TimeZoneInfo]::ConvertTimeToUtc($unspecified, $tz)
    return @{
        matches = $true
        publicationDate = $localDate.Date
        cycleStartUtc = $cycleUtc
        reason = 'scheduled_slot'
        resolvedTimeZoneId = [string]$tz.Id
    }
}

function _SSPub-DocProp([object]$Doc, [string[]]$Names) {
    if ($null -eq $Doc) { return $null }
    foreach ($n in @($Names)) {
        try {
            if ($Doc -is [hashtable] -and $Doc.ContainsKey($n) -and $null -ne $Doc[$n] -and -not (_SSPub-IsBlank $Doc[$n])) { return $Doc[$n] }
            if ($Doc.PSObject -and $Doc.PSObject.Properties[$n] -and $null -ne $Doc.$n -and -not (_SSPub-IsBlank ([string]$Doc.$n))) { return $Doc.$n }
        } catch { }
    }
    return $null
}

function _SSPub-DocName([object]$Doc) {
    $n = _SSPub-DocProp -Doc $Doc -Names @('Name', 'DocumentName')
    if ($n) { return [string]$n }
    return ''
}

function _SSPub-DocLastWrite([object]$Doc) {
    $raw = _SSPub-DocProp -Doc $Doc -Names @(
        'FileUpdateDateUtc', 'FileUpdatedDateUtc', 'FileUpdateDate', 'FileUpdatedDate',
        'DocumentUpdateDate', 'LastWriteUtc', 'LastWriteTimeUtc'
    )
    if ($null -eq $raw) { return $null }
    if ($raw -is [datetime]) { return [datetime]$raw }
    try { return [datetime]$raw } catch { return $null }
}

function _SSPub-ToUtc([object]$Value, $TimeZone) {
    if ($null -eq $Value) { return $null }
    $dt = [datetime]$Value
    if ($dt.Kind -eq [DateTimeKind]::Utc) { return $dt }
    if ($dt.Kind -eq [DateTimeKind]::Local) { return $dt.ToUniversalTime() }
    if ($TimeZone) {
        $unspecified = [datetime]::SpecifyKind($dt, [DateTimeKind]::Unspecified)
        return [TimeZoneInfo]::ConvertTimeToUtc($unspecified, $TimeZone)
    }
    return [datetime]::SpecifyKind($dt, [DateTimeKind]::Utc)
}

function _SSPub-FindDocument {
    param($DocumentGateway, [string]$FolderPath, [string]$DocumentName)
    if ($DocumentGateway) {
        return & $DocumentGateway.FindDocument $FolderPath $DocumentName
    }
    $cmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return @{ _missingCmdlet = $true }
    }
    foreach ($tp in @(_SSPub-FolderTryPaths -Path $FolderPath)) {
        try {
            $doc = Get-PWDocumentsBySearch -FolderPath $tp -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($doc) { return $doc }
        } catch { }
    }
    return $null
}

function _SSPub-ExportDocument {
    param($DocumentGateway, $Document, [string]$TargetFolder)
    if (-not (Test-Path -LiteralPath $TargetFolder)) {
        New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
    }
    if ($DocumentGateway) {
        $path = & $DocumentGateway.ExportDocument $Document $TargetFolder
        if (_SSPub-IsBlank $path) { throw 'Gateway export returned no path.' }
        return [string]$path
    }
    $exported = Export-StatusSetPdfToFolder -InputDocument $Document -TargetFolder $TargetFolder
    if (-not $exported.IsSuccess) { throw $exported.Message }
    return [string]$exported.Data.localPath
}

function _SSPub-CreateDocument {
    param($DocumentGateway, [string]$FolderPath, [string]$FilePath, [string]$DocumentName)
    if ($DocumentGateway) {
        & $DocumentGateway.CreateDocument $FolderPath $FilePath $DocumentName
        return
    }
    $cmd = Get-Command -Name 'New-PWDocument' -ErrorAction SilentlyContinue
    if (-not $cmd) { throw 'New-PWDocument is not available in this session.' }
    $created = $false
    $lastError = 'No resolved ProjectWise folder accepted the document.'
    foreach ($tp in @(_SSPub-FolderTryPaths -Path $FolderPath)) {
        if (-not (_SSPub-PwFolderExists -Path $tp)) { continue }
        try {
            $createdDoc = New-PWDocument -FolderPath $tp -FilePath $FilePath -DocumentName $DocumentName -ErrorAction Stop
            if ($createdDoc) {
                $created = $true
                break
            }
            $lastError = "New-PWDocument returned no document for folder '$tp'."
        } catch {
            $lastError = [string]$_.Exception.Message
        }
    }
    if (-not $created) { throw $lastError }
}

function _SSPub-ReadPdfHeaderOk([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        if ($fs.Length -lt 5) { return $false }
        $buf = New-Object byte[] 5
        [void]$fs.Read($buf, 0, 5)
        $head = [System.Text.Encoding]::ASCII.GetString($buf, 0, 4)
        return ($head -eq '%PDF')
    } catch {
        return $false
    } finally {
        if ($fs) { $fs.Dispose() }
    }
}

function _SSPub-FileLength([string]$Path) {
    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $item) { return [int64]0 }
    return [int64]$item.Length
}

function _SSPub-NewSegment([hashtable]$Set) {
    return @{
        id = [string]$Set.id
        sourceStatus = 'MISSING'
        sourceName = [string]$Set.documentName
        sourceFolder = [string]$Set.sourceFolder
        sourceLastWriteUtc = $null
        sourceSize = [int64]0
        sourceHash = ''
        validation = 'FAILED'
        validationDetail = ''
        provenance = ''
        newlyGenerated = $false
        destinationFolder = [string]$Set.destinationFolder
        destinationName = ''
        conflict = $false
        copy = 'NOT_RUN'
        verification = 'NOT_RUN'
        canPublish = $false
        ok = $false
        deferred = $false
        code = ''
        message = ''
    }
}

function _SSPub-ManifestGeneratedUtc {
    param([hashtable]$Config, [string]$SourceFolder, $TimeZone)
    if (-not $Config -or -not $Config.ContainsKey('statusSet')) { return $null }
    $ss = _SSPub-AsHashtable $Config['statusSet']
    if (-not $ss -or -not $ss.ContainsKey('localRoot') -or (_SSPub-IsBlank $ss['localRoot'])) { return $null }
    if (-not (Get-Command -Name 'Get-StatusSetWorkspaceDirectory' -ErrorAction SilentlyContinue)) { return $null }
    $display = _SSPub-DocumentsDisplayPath -Path $SourceFolder
    $ws = Get-StatusSetWorkspaceDirectory -LocalRoot ([string]$ss['localRoot']) -SheetsFolderPath $display
    if (-not $ws.IsSuccess) { return $null }
    $manifestName = '_statusset.manifest.json'
    if ($ss.ContainsKey('manifestFileName') -and -not (_SSPub-IsBlank $ss['manifestFileName'])) {
        $manifestName = [string]$ss['manifestFileName']
    }
    $manifestPath = Get-StatusSetManifestPath -WorkspaceDir ([string]$ws.Data.workspaceDir) -ManifestFileName $manifestName
    $read = Read-StatusSetManifestFile -Path $manifestPath
    if (-not $read.IsSuccess -or -not $read.Data.manifest) { return $null }
    $man = $read.Data.manifest
    $raw = $null
    if ($man -is [hashtable] -and $man.ContainsKey('generatedAtUtc')) { $raw = $man['generatedAtUtc'] }
    elseif ($man.PSObject -and $man.PSObject.Properties['generatedAtUtc']) { $raw = $man.generatedAtUtc }
    if (_SSPub-IsBlank $raw) { return $null }
    try { return _SSPub-ToUtc -Value ([datetime]$raw) -TimeZone $TimeZone } catch { return $null }
}

function _SSPub-LiveConfirmedCurrent {
    param([hashtable]$Config, [string]$SourceFolder)
    if (-not (Get-Command -Name 'Get-StatusSetPWFolderState' -ErrorAction SilentlyContinue)) { return $false }
    if (-not (Get-Command -Name 'Test-StatusSetWatcherShouldEnqueue' -ErrorAction SilentlyContinue)) { return $false }
    try {
        $display = _SSPub-DocumentsDisplayPath -Path $SourceFolder
        $pwPath = $display
        if ($pwPath -match '^(?i)Documents\\') { $pwPath = $pwPath.Substring('Documents\'.Length) }
        $state = Get-StatusSetPWFolderState -FolderPath $pwPath -OneLevelDeep $false
        $gate = Test-StatusSetWatcherShouldEnqueue -Config $Config -SourceFolder $display -FolderState $state
        if ($gate.IsSuccess -and -not [bool]$gate.Data.shouldEnqueue -and [string]$gate.Data.gateReason -eq 'ALREADY_CURRENT') {
            return $true
        }
    } catch { }
    return $false
}

function _SSPub-InFlight {
    param([hashtable]$Config, [string]$SourceFolder, [string]$SetId, [string[]]$InFlightSetIds, [bool]$CheckInFlight)
    if (-not $CheckInFlight) { return $false }
    if ($InFlightSetIds -and (@($InFlightSetIds) -contains $SetId)) { return $true }
    if (-not (Get-Command -Name 'Test-QCStatusSetJobInFlight' -ErrorAction SilentlyContinue)) { return $false }
    try {
        $res = Test-QCStatusSetJobInFlight -Config $Config -SourceFolder $SourceFolder
        if ($res.IsSuccess -and [bool]$res.Data.inFlight) { return $true }
    } catch { }
    return $false
}

function Invoke-StatusSetPublicationSourceReconcile {
    <#
    .SYNOPSIS
    Runs the existing STATUS_SET_GEN processor for configured source folders.
    .DESCRIPTION
    Does not mark full-scan schedule slots complete and does not move the audit watermark.
    A full Watch-QCTrigger reconciliation cannot be started from here without those side effects.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ProjectKey,
        [scriptblock]$ReconcileCommand
    )

    $settingsRes = Get-StatusSetPublicationSettings -Config $Config
    if (-not $settingsRes.IsSuccess) { return $settingsRes }
    $project = _SSPub-FindProject -Settings $settingsRes.Data -ProjectKey $ProjectKey
    if (-not $project) {
        return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_CONFIG' -Message "Project '$ProjectKey' has no status set publication configuration." -Data @{
            projectKey = $ProjectKey
        }
    }
    if (-not $ReconcileCommand) {
        $ReconcileCommand = {
            param($Job, $ReconcileConfig)
            Invoke-StatusSetNativeJob -Job $Job -Config $ReconcileConfig
        }
    }

    $hints = @{}
    $rows = New-Object System.Collections.Generic.List[hashtable]
    foreach ($set in @($project.sets)) {
        $problems = @(_SSPub-ValidateSet -Set $set)
        if ($problems.Count -gt 0) {
            $row = @{
                id = [string]$set.id
                ok = $false
                failed = $true
                newlyGenerated = $false
                kind = 'failed'
                reconcileCode = 'STATUS_SET_PUBLISH_CONFIG'
                reconcileMessage = ($problems -join '; ')
            }
            $hints[[string]$set.id] = $row
            [void]$rows.Add($row)
            continue
        }
        $job = @{
            id = ('manual-statusset-publication-' + [string]$set.id)
            type = 'STATUS_SET_GEN'
            sourceFolder = _SSPub-DocumentsDisplayPath -Path ([string]$set.sourceFolder)
            metadata = @{
                runKind = 'Manual'
                manualStatusSetPublication = $true
                setId = [string]$set.id
            }
        }
        $result = $null
        try {
            $result = & $ReconcileCommand $job $Config
        } catch {
            $result = New-QCFailureResult -Code 'STATUS_SET_PUBLISH_RECONCILE_FAILED' -Message $_.Exception.Message -Data @{}
        }
        $code = if ($result -and $result.Code) { [string]$result.Code } else { 'STATUS_SET_PUBLISH_RECONCILE_FAILED' }
        $message = if ($result -and $result.Message) { [string]$result.Message } else { 'Reconcile returned no result.' }
        $generated = $false
        $kind = 'failed'
        $failed = $true
        if ($result -and $result.IsSuccess -and $code -in @('STATUS_SET_OK', 'STATUS_SET_RECONCILE_CREATED', 'STATUS_SET_RECONCILE_UPDATED')) {
            $generated = $true
            $kind = 'refreshed'
            $failed = $false
        } elseif ($result -and $result.IsSuccess -and $code -in @('STATUS_SET_UP_TO_DATE', 'STATUS_SET_RECONCILE_IN_SYNC', 'STATUS_SET_DRYRUN', 'STATUS_SET_DRYRUN_REBUILD')) {
            $generated = $false
            $kind = 'confirmed_current'
            $failed = $false
        }
        $row = @{
            id = [string]$set.id
            ok = -not $failed
            failed = $failed
            newlyGenerated = $generated
            kind = $kind
            reconcileCode = $code
            reconcileMessage = $message
        }
        $hints[[string]$set.id] = $row
        [void]$rows.Add($row)
    }

    $anyFail = @($rows | Where-Object { $_.failed }).Count -gt 0
    return New-QCSuccessResult -Code $(if ($anyFail) { 'STATUS_SET_PUBLISH_RECONCILE_PARTIAL' } else { 'STATUS_SET_PUBLISH_RECONCILE_DONE' }) -Message 'Manual status-set reconcile finished without changing the reconciliation schedule.' -Data @{
        projectKey = $ProjectKey
        runKind = 'Manual'
        scheduleStateModified = $false
        hints = $hints
        sets = @($rows)
    }
}

function Invoke-StatusSetPublication {
    <#
    .SYNOPSIS
    Validates and, in Publish mode, copies configured Status Set PDFs to their archive names.
    .DESCRIPTION
    Preview performs no ProjectWise creates. Identical destination content is already published.
    Different content under the same name is a conflict and is not overwritten.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$ProjectKey,
        [Parameter(Mandatory)][datetime]$AsOfDate,
        [ValidateSet('Preview', 'Publish')]
        [string]$Mode = 'Preview',
        [switch]$UseLatestExisting,
        [ValidateSet('Manual', 'Scheduled')]
        [string]$RunKind = 'Manual',
        [Nullable[datetime]]$CycleStartUtc,
        [hashtable]$ProvenanceHints,
        [bool]$CheckInFlight = $false,
        [string[]]$InFlightSetIds = @(),
        [hashtable]$DocumentGateway,
        [scriptblock]$LogCallback
    )

    $settingsRes = Get-StatusSetPublicationSettings -Config $Config
    if (-not $settingsRes.IsSuccess) { return $settingsRes }
    $settings = $settingsRes.Data
    $project = _SSPub-FindProject -Settings $settings -ProjectKey $ProjectKey
    if (-not $project) {
        return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_CONFIG' -Message "Project '$ProjectKey' has no status set publication configuration." -Data @{
            projectKey = $ProjectKey
            mode = $Mode
            overall = 'FAILURE'
            segments = @()
        }
    }
    if (-not $project.sets -or @($project.sets).Count -eq 0) {
        return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_CONFIG' -Message "Project '$ProjectKey' has no status sets configured." -Data @{
            projectKey = $ProjectKey
            mode = $Mode
            overall = 'FAILURE'
            segments = @()
        }
    }

    $tzRes = Resolve-StatusSetPublicationTimeZone -TimeZoneId ([string]$settings.timeZoneId)
    $tz = if ($tzRes.IsSuccess) { $tzRes.Data.timeZone } else { $null }
    $requireProvenance = -not $UseLatestExisting.IsPresent
    $publicationDate = $AsOfDate.Date
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('QC_StatusSetPublication_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    $dryRun = $false
    if ($Config.ContainsKey('dryRun') -and $null -ne $Config['dryRun']) {
        try { $dryRun = [bool]$Config['dryRun'] } catch { $dryRun = $false }
    }

    _SSPub-Log -LogCallback $LogCallback -Level 'Information' -Code 'STATUS_SET_PUBLISH_START' -Message 'Status set publication started.' -Data @{
        projectKey = $ProjectKey
        mode = $Mode
        runKind = $RunKind
        asOfDate = $publicationDate.ToString('yyyy-MM-dd')
        useLatestExisting = [bool]$UseLatestExisting.IsPresent
        requireProvenance = $requireProvenance
    }

    $segments = New-Object System.Collections.Generic.List[hashtable]
    try {
        foreach ($set in @($project.sets)) {
            $seg = _SSPub-NewSegment -Set $set
            $problems = @(_SSPub-ValidateSet -Set $set)
            if ($problems.Count -gt 0) {
                $seg.validation = 'FAILED'
                $seg.validationDetail = ($problems -join '; ')
                $seg.code = 'STATUS_SET_PUBLISH_CONFIG'
                $seg.message = $seg.validationDetail
                [void]$segments.Add($seg)
                continue
            }

            $destName = Format-StatusSetPublicationFileName -Pattern ([string]$set.fileNamePattern) -AsOfDate $publicationDate -SetId ([string]$set.id)
            $seg.destinationName = $destName
            $hint = $null
            if ($ProvenanceHints -and $ProvenanceHints.ContainsKey([string]$set.id)) {
                $hint = _SSPub-AsHashtable $ProvenanceHints[[string]$set.id]
            }

            if ($hint -and [bool]$hint['failed'] -and -not $UseLatestExisting.IsPresent) {
                $seg.validation = 'FAILED'
                $seg.provenance = 'failed'
                $seg.newlyGenerated = $false
                $seg.validationDetail = [string]$hint['reconcileMessage']
                $seg.code = [string]$hint['reconcileCode']
                $seg.message = "Reconcile did not produce a publishable status set: $($seg.validationDetail)"
                [void]$segments.Add($seg)
                continue
            }

            if (_SSPub-InFlight -Config $Config -SourceFolder ([string]$set.sourceFolder) -SetId ([string]$set.id) -InFlightSetIds $InFlightSetIds -CheckInFlight $CheckInFlight) {
                $seg.sourceStatus = 'PENDING'
                $seg.validation = 'FAILED'
                $seg.deferred = $true
                $seg.provenance = 'deferred'
                $seg.newlyGenerated = $false
                $seg.code = 'STATUS_SET_PUBLISH_DEFERRED'
                $seg.message = 'STATUS_SET_GEN is still pending or running for this folder.'
                $seg.validationDetail = $seg.message
                [void]$segments.Add($seg)
                continue
            }

            $sourceDoc = _SSPub-FindDocument -DocumentGateway $DocumentGateway -FolderPath ([string]$set.sourceFolder) -DocumentName ([string]$set.documentName)
            if ($sourceDoc -is [hashtable] -and $sourceDoc.ContainsKey('_missingCmdlet')) {
                $seg.sourceStatus = 'MISSING'
                $seg.validationDetail = 'Get-PWDocumentsBySearch is not available.'
                $seg.code = 'STATUS_SET_PUBLISH_PW_CMD_MISSING'
                $seg.message = $seg.validationDetail
                [void]$segments.Add($seg)
                continue
            }
            if (-not $sourceDoc) {
                $seg.sourceStatus = 'MISSING'
                $seg.validationDetail = "Source document '$($set.documentName)' was not found."
                $seg.code = 'STATUS_SET_PUBLISH_SOURCE_MISSING'
                $seg.message = $seg.validationDetail
                [void]$segments.Add($seg)
                continue
            }

            $seg.sourceStatus = 'FOUND'
            $seg.sourceName = _SSPub-DocName -Doc $sourceDoc
            if (_SSPub-IsBlank $seg.sourceName) { $seg.sourceName = [string]$set.documentName }
            $srcWrite = _SSPub-DocLastWrite -Doc $sourceDoc
            $seg.sourceLastWriteUtc = _SSPub-ToUtc -Value $srcWrite -TimeZone $tz
            $sizeProp = _SSPub-DocProp -Doc $sourceDoc -Names @('FileSize', 'Size')
            if ($null -ne $sizeProp) {
                try { $seg.sourceSize = [int64]$sizeProp } catch { $seg.sourceSize = [int64]0 }
            }

            $sourceExportDir = Join-Path $tempRoot ('src_' + [string]$set.id)
            $sourceLocal = $null
            try {
                $sourceLocal = _SSPub-ExportDocument -DocumentGateway $DocumentGateway -Document $sourceDoc -TargetFolder $sourceExportDir
            } catch {
                $seg.sourceStatus = 'INVALID'
                $seg.validationDetail = ('Source export failed: ' + $_.Exception.Message)
                $seg.code = 'STATUS_SET_PUBLISH_SOURCE_EXPORT_FAILED'
                $seg.message = $seg.validationDetail
                [void]$segments.Add($seg)
                continue
            }
            $seg.sourceSize = _SSPub-FileLength -Path $sourceLocal
            $headerOk = _SSPub-ReadPdfHeaderOk -Path $sourceLocal
            if ($seg.sourceSize -le 0 -or -not $headerOk) {
                $seg.sourceStatus = 'INVALID'
                $seg.validation = 'FAILED'
                $seg.validationDetail = 'Source is missing, empty, or not a PDF.'
                $seg.code = 'STATUS_SET_PUBLISH_SOURCE_INVALID'
                $seg.message = $seg.validationDetail
                [void]$segments.Add($seg)
                continue
            }
            $sourceHash = _SSPub-Sha256File -Path $sourceLocal
            $seg.sourceHash = $sourceHash

            $provenanceOk = $true
            $provenance = 'latest_existing'
            $newly = $false
            if ($UseLatestExisting.IsPresent) {
                $provenance = 'latest_existing'
                $newly = $false
            } elseif ($hint -and -not [bool]$hint['failed'] -and -not (_SSPub-IsBlank $hint['kind'])) {
                $provenance = [string]$hint['kind']
                $newly = [bool]$hint['newlyGenerated']
            } elseif (-not $requireProvenance) {
                $provenance = 'latest_existing'
            } else {
                if ($null -eq $CycleStartUtc) {
                    $provenanceOk = $false
                    $provenance = 'failed'
                    $seg.validationDetail = 'Cycle provenance requires a reconciliation cycle. Pass -UseLatestExisting or -RunReconciliation for a manual run.'
                } else {
                    $cycleUtc = _SSPub-ToUtc -Value $CycleStartUtc -TimeZone $tz
                    $srcUtc = $seg.sourceLastWriteUtc
                    $manUtc = $null
                    if (-not $DocumentGateway) {
                        $manUtc = _SSPub-ManifestGeneratedUtc -Config $Config -SourceFolder ([string]$set.sourceFolder) -TimeZone $tz
                    }
                    if ($srcUtc -and $srcUtc -ge $cycleUtc) {
                        $provenance = 'refreshed'
                        $newly = $true
                    } elseif ($manUtc -and $manUtc -ge $cycleUtc) {
                        $provenance = 'refreshed'
                        $newly = $true
                    } elseif ((-not $DocumentGateway) -and (_SSPub-LiveConfirmedCurrent -Config $Config -SourceFolder ([string]$set.sourceFolder))) {
                        $provenance = 'confirmed_current'
                        $newly = $false
                    } else {
                        $provenanceOk = $false
                        $provenance = 'failed'
                        $shown = if ($srcUtc) { $srcUtc.ToString('o') } else { 'unknown' }
                        $seg.validationDetail = "Source was not generated or confirmed during this reconciliation cycle (source date $shown)."
                    }
                }
            }
            $seg.provenance = $provenance
            $seg.newlyGenerated = $newly
            if (-not $provenanceOk) {
                $seg.validation = 'FAILED'
                $seg.code = 'STATUS_SET_PUBLISH_PROVENANCE'
                $seg.message = $seg.validationDetail
                [void]$segments.Add($seg)
                continue
            }

            $seg.validation = 'PASSED'
            $seg.validationDetail = if ($UseLatestExisting.IsPresent) {
                $shown = if ($seg.sourceLastWriteUtc) { $seg.sourceLastWriteUtc.ToString('o') } else { 'unknown' }
                "UseLatestExisting bypass; source date $shown. Not marked newly generated."
            } elseif ($newly) {
                'Source was generated or refreshed for this cycle.'
            } else {
                'Source confirmed current for this cycle and was not newly generated.'
            }

            $destDoc = _SSPub-FindDocument -DocumentGateway $DocumentGateway -FolderPath ([string]$set.destinationFolder) -DocumentName $destName
            if ($destDoc -is [hashtable] -and $destDoc.ContainsKey('_missingCmdlet')) {
                $seg.code = 'STATUS_SET_PUBLISH_PW_CMD_MISSING'
                $seg.message = 'Get-PWDocumentsBySearch is not available.'
                $seg.validation = 'FAILED'
                [void]$segments.Add($seg)
                continue
            }
            if ($destDoc) {
                $destExportDir = Join-Path $tempRoot ('dest_' + [string]$set.id)
                try {
                    $destLocal = _SSPub-ExportDocument -DocumentGateway $DocumentGateway -Document $destDoc -TargetFolder $destExportDir
                    $destHash = _SSPub-Sha256File -Path $destLocal
                } catch {
                    $seg.validation = 'FAILED'
                    $seg.verification = 'FAILED'
                    $seg.code = 'STATUS_SET_PUBLISH_DEST_EXPORT_FAILED'
                    $seg.message = ('Could not read the existing destination document: ' + $_.Exception.Message)
                    [void]$segments.Add($seg)
                    continue
                }
                if ($destHash -eq $sourceHash) {
                    $seg.copy = 'ALREADY_PUBLISHED'
                    $seg.verification = 'PASSED'
                    $seg.canPublish = $true
                    $seg.ok = $true
                    $seg.conflict = $false
                    $seg.code = 'STATUS_SET_PUBLISH_ALREADY'
                    $seg.message = 'Destination already matches the source. No copy performed.'
                    [void]$segments.Add($seg)
                    continue
                }
                $seg.conflict = $true
                $seg.copy = 'NOT_RUN'
                $seg.verification = 'FAILED'
                $seg.canPublish = $false
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_CONFLICT'
                $seg.message = "Destination '$destName' exists with different content and was not overwritten."
                [void]$segments.Add($seg)
                continue
            }

            $seg.canPublish = $true
            if ($Mode -eq 'Preview' -or $dryRun) {
                $seg.copy = 'NOT_RUN'
                $seg.verification = 'NOT_RUN'
                $seg.ok = $true
                $seg.code = if ($dryRun -and $Mode -eq 'Publish') { 'STATUS_SET_PUBLISH_DRYRUN' } else { 'STATUS_SET_PUBLISH_PREVIEW' }
                $seg.message = if ($dryRun -and $Mode -eq 'Publish') { 'dryRun is true. No ProjectWise copy was made.' } else { 'Preview only. No ProjectWise copy was made.' }
                [void]$segments.Add($seg)
                continue
            }

            try {
                _SSPub-CreateDocument -DocumentGateway $DocumentGateway -FolderPath ([string]$set.destinationFolder) -FilePath $sourceLocal -DocumentName $destName
                $seg.copy = 'SUCCESS'
            } catch {
                $seg.copy = 'FAILED'
                $seg.verification = 'NOT_RUN'
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_COPY_FAILED'
                $seg.message = ('Copy failed: ' + $_.Exception.Message)
                [void]$segments.Add($seg)
                continue
            }

            $sourceAfter = _SSPub-FindDocument -DocumentGateway $DocumentGateway -FolderPath ([string]$set.sourceFolder) -DocumentName ([string]$set.documentName)
            $sourceStill = $null
            try {
                $sourceStillDir = Join-Path $tempRoot ('src2_' + [string]$set.id)
                $sourceStill = _SSPub-ExportDocument -DocumentGateway $DocumentGateway -Document $sourceAfter -TargetFolder $sourceStillDir
            } catch {
                $seg.verification = 'FAILED'
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_SOURCE_CHANGED'
                $seg.message = 'Could not re-read the source after copy.'
                [void]$segments.Add($seg)
                continue
            }
            $sourceHashAfter = _SSPub-Sha256File -Path $sourceStill
            if ($sourceHashAfter -ne $sourceHash) {
                $seg.verification = 'FAILED'
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_SOURCE_CHANGED'
                $seg.message = 'Source PDF changed during publication. The source must remain unmodified.'
                [void]$segments.Add($seg)
                continue
            }

            $copied = _SSPub-FindDocument -DocumentGateway $DocumentGateway -FolderPath ([string]$set.destinationFolder) -DocumentName $destName
            if (-not $copied) {
                $seg.verification = 'FAILED'
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_VERIFY_MISSING'
                $seg.message = 'Copy was requested but the destination document was not found.'
                [void]$segments.Add($seg)
                continue
            }
            try {
                $verifyDir = Join-Path $tempRoot ('verify_' + [string]$set.id)
                $verifyLocal = _SSPub-ExportDocument -DocumentGateway $DocumentGateway -Document $copied -TargetFolder $verifyDir
                $verifyHash = _SSPub-Sha256File -Path $verifyLocal
                $verifySize = _SSPub-FileLength -Path $verifyLocal
            } catch {
                $seg.verification = 'FAILED'
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_VERIFY_FAILED'
                $seg.message = ('Destination could not be read back: ' + $_.Exception.Message)
                [void]$segments.Add($seg)
                continue
            }
            if ($verifyHash -ne $sourceHash -or $verifySize -ne $seg.sourceSize) {
                $seg.verification = 'FAILED'
                $seg.conflict = $true
                $seg.ok = $false
                $seg.code = 'STATUS_SET_PUBLISH_VERIFY_MISMATCH'
                $seg.message = 'Destination exists but its content does not match the source. It was not overwritten.'
                [void]$segments.Add($seg)
                continue
            }

            $seg.verification = 'PASSED'
            $seg.ok = $true
            $seg.code = 'STATUS_SET_PUBLISH_COPIED'
            $seg.message = 'Destination copied and verified.'
            [void]$segments.Add($seg)
        }
    } finally {
        try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }

    $anyDeferred = @($segments | Where-Object { $_.deferred }).Count -gt 0
    $anyFailed = @($segments | Where-Object { -not $_.ok -and -not $_.deferred }).Count -gt 0
    $overall = 'SUCCESS'
    $code = if ($Mode -eq 'Preview') { 'STATUS_SET_PUBLISH_PREVIEW_OK' } else { 'STATUS_SET_PUBLISH_OK' }
    $ok = $true
    if ($anyFailed) {
        $overall = 'FAILURE'
        $code = 'STATUS_SET_PUBLISH_FAILED'
        $ok = $false
    } elseif ($anyDeferred) {
        $overall = 'DEFERRED'
        $code = 'STATUS_SET_PUBLISH_DEFERRED'
        $ok = $true
    } elseif ($dryRun -and $Mode -eq 'Publish') {
        $overall = 'FAILURE'
        $code = 'STATUS_SET_PUBLISH_DRYRUN'
        $ok = $false
    }

    $level = if ($ok -and $overall -ne 'DEFERRED') { 'Information' } elseif ($overall -eq 'DEFERRED') { 'Warning' } else { 'Error' }
    _SSPub-Log -LogCallback $LogCallback -Level $level -Code $code -Message "Status set publication $overall." -Data @{
        projectKey = $ProjectKey
        mode = $Mode
        runKind = $RunKind
        overall = $overall
        segmentCount = @($segments).Count
    }

    $payload = @{
        projectKey = $ProjectKey
        mode = $Mode
        runKind = $RunKind
        asOfDate = $publicationDate.ToString('yyyy-MM-dd')
        useLatestExisting = [bool]$UseLatestExisting.IsPresent
        scheduledEnabled = [bool]$settings.enabled
        overall = $overall
        segments = @($segments)
    }
    if ($ok) {
        return New-QCSuccessResult -Code $code -Message "Status set publication $overall." -Data $payload
    }
    return New-QCFailureResult -Code $code -Message "Status set publication $overall." -Data $payload
}

function _SSPub-PendingPath([string]$QueueRoot) {
    if (_SSPub-IsBlank $QueueRoot) { return $null }
    return Join-Path (Join-Path $QueueRoot '_watcher') $script:SSPubPendingFileName
}

function _SSPub-ReadPending([string]$Path) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if (_SSPub-IsBlank $raw) { return $null }
        return _SSPub-AsHashtable ($raw | ConvertFrom-Json)
    } catch { return $null }
}

function _SSPub-WritePending([string]$Path, [hashtable]$Doc) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $json = $Doc | ConvertTo-Json -Depth 6
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Invoke-StatusSetScheduledPublication {
    <#
    .SYNOPSIS
    Publishes after a completed Monday 06:00 full-scan slot, and retries a pending Monday publish on later ticks.
    .DESCRIPTION
    No-ops when statusSetPublication.enabled is false. Does not mark reconciliation slots complete.
    Failures are returned and logged; callers must not let them abort unrelated reconciliation work.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$QueueRoot = '',
        [string]$CompletedSlotKey = '',
        [scriptblock]$LogCallback,
        [hashtable]$DocumentGateway,
        [string[]]$InFlightSetIds = @()
    )

    try {
        $settingsRes = Get-StatusSetPublicationSettings -Config $Config
        if (-not $settingsRes.IsSuccess) { return $settingsRes }
        $settings = $settingsRes.Data
        if (-not [bool]$settings.enabled) {
            $slot = Test-StatusSetPublicationSlot -Settings $settings -SlotKey $CompletedSlotKey
            if ($slot.matches) {
                _SSPub-Log -LogCallback $LogCallback -Level 'Information' -Code 'STATUS_SET_PUBLISH_DISABLED' -Message 'Monday reconciliation finished. Status set publication is not enabled, so no archive copy was made.' -Data @{
                    completedSlotKey = $CompletedSlotKey
                }
            }
            return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_DISABLED' -Message 'Scheduled publication is disabled.' -Data @{
                overall = 'SKIPPED'
                scheduledEnabled = $false
                scheduleStateModified = $false
            }
        }
        if (-not $settings.projects -or @($settings.projects | Where-Object { [bool]$_.enabled -and @($_.sets).Count -gt 0 }).Count -eq 0) {
            return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_NONE' -Message 'No enabled projects have status sets configured.' -Data @{
                overall = 'SKIPPED'
                scheduleStateModified = $false
            }
        }

        $pendingPath = _SSPub-PendingPath -QueueRoot $QueueRoot
        $pending = _SSPub-ReadPending -Path $pendingPath
        $slot = Test-StatusSetPublicationSlot -Settings $settings -SlotKey $CompletedSlotKey
        $active = $null
        if ($slot.matches) {
            $active = @{
                slotKey = $CompletedSlotKey
                publicationDate = ([datetime]$slot.publicationDate).ToString('yyyy-MM-dd')
                cycleStartUtc = ([datetime]$slot.cycleStartUtc).ToString('o')
                createdUtc = (Get-Date).ToUniversalTime().ToString('o')
                attempts = 0
            }
        } elseif ($pending) {
            $created = $null
            try { if ($pending.ContainsKey('createdUtc')) { $created = [datetime]$pending['createdUtc'] } } catch { $created = $null }
            $expired = $false
            if ($created) {
                $age = (Get-Date).ToUniversalTime() - $created.ToUniversalTime()
                if ($age.TotalHours -gt $script:SSPubPendingMaxAgeHours) { $expired = $true }
            }
            if ($expired) {
                if ($pendingPath -and (Test-Path -LiteralPath $pendingPath)) {
                    Remove-Item -LiteralPath $pendingPath -Force -ErrorAction SilentlyContinue
                }
                _SSPub-Log -LogCallback $LogCallback -Level 'Warning' -Code 'STATUS_SET_PUBLISH_EXPIRED' -Message 'Pending status set publication expired and was not retried.' -Data @{
                    slotKey = [string]$pending['slotKey']
                }
                return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_EXPIRED' -Message 'Pending publication expired.' -Data @{
                    overall = 'FAILURE'
                    scheduleStateModified = $false
                }
            }
            $active = $pending
        } else {
            return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_NOT_DUE' -Message 'This reconciliation slot does not publish status sets.' -Data @{
                overall = 'SKIPPED'
                reason = [string]$slot.reason
                scheduleStateModified = $false
            }
        }

        $asOf = [datetime]::ParseExact([string]$active.publicationDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        $cycle = [datetime]$active.cycleStartUtc
        $attempts = 0
        try { $attempts = [int]$active.attempts } catch { $attempts = 0 }
        $attempts++

        $projectResults = New-Object System.Collections.ArrayList
        $anyDeferred = $false
        $anyFailed = $false
        foreach ($project in @($settings.projects)) {
            if (-not [bool]$project.enabled) { continue }
            if (-not $project.sets -or @($project.sets).Count -eq 0) { continue }
            $one = Invoke-StatusSetPublication -Config $Config -ProjectKey ([string]$project.projectKey) -AsOfDate $asOf -Mode Publish -RunKind Scheduled -CycleStartUtc $cycle -CheckInFlight:$true -InFlightSetIds $InFlightSetIds -DocumentGateway $DocumentGateway -LogCallback $LogCallback
            [void]$projectResults.Add($one)
            if ($one.Data -and [string]$one.Data.overall -eq 'DEFERRED') { $anyDeferred = $true }
            if (-not $one.IsSuccess) { $anyFailed = $true }
        }

        if ($anyDeferred -and $pendingPath) {
            _SSPub-WritePending -Path $pendingPath -Doc @{
                slotKey = [string]$active.slotKey
                publicationDate = [string]$active.publicationDate
                cycleStartUtc = [string]$active.cycleStartUtc
                createdUtc = [string]$active.createdUtc
                attempts = $attempts
                lastResult = 'DEFERRED'
            }
            _SSPub-Log -LogCallback $LogCallback -Level 'Warning' -Code 'STATUS_SET_PUBLISH_DEFERRED' -Message 'Status set publication deferred until STATUS_SET_GEN jobs finish. Reconciliation schedule was not changed.' -Data @{
                slotKey = [string]$active.slotKey
                attempts = $attempts
            }
            return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_DEFERRED' -Message 'Publication deferred.' -Data @{
                overall = 'DEFERRED'
                scheduleStateModified = $false
                attempts = $attempts
                projects = @($projectResults)
            }
        }

        if ($pendingPath -and (Test-Path -LiteralPath $pendingPath)) {
            Remove-Item -LiteralPath $pendingPath -Force -ErrorAction SilentlyContinue
        }

        if ($anyFailed) {
            return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_FAILED' -Message 'One or more status set publications failed.' -Data @{
                overall = 'FAILURE'
                scheduleStateModified = $false
                projects = @($projectResults)
            }
        }
        return New-QCSuccessResult -Code 'STATUS_SET_PUBLISH_OK' -Message 'Scheduled status set publication finished.' -Data @{
            overall = 'SUCCESS'
            scheduleStateModified = $false
            projects = @($projectResults)
        }
    } catch {
        _SSPub-Log -LogCallback $LogCallback -Level 'Error' -Code 'STATUS_SET_PUBLISH_FAILED' -Message ('Scheduled publication threw: ' + $_.Exception.Message) -Data @{}
        return New-QCFailureResult -Code 'STATUS_SET_PUBLISH_FAILED' -Message $_.Exception.Message -Data @{
            overall = 'FAILURE'
            scheduleStateModified = $false
        }
    }
}

Export-ModuleMember -Function @(
    'Resolve-StatusSetPublicationTimeZone',
    'ConvertTo-StatusSetPublicationFolderKey',
    'Get-StatusSetPublicationSettings',
    'Format-StatusSetPublicationFileName',
    'Test-StatusSetPublicationSlot',
    'Invoke-StatusSetPublicationSourceReconcile',
    'Invoke-StatusSetPublication',
    'Invoke-StatusSetScheduledPublication'
)
