# PW.Connection.psm1
# Responsibility: ProjectWise connection port (read-only safe by default).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

function _PWC-TestPWDiscoveryCmdlets {
    [CmdletBinding()]
    param()

    # pwps_dab discovery/listing can appear independently of pwps connection cmdlets.
    # Either command is enough to prove the current runspace still has document/folder discovery.
    return [bool]((Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue) -or (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue))
}

function _PWC-TryImportPWModules {
    <#
    .SYNOPSIS
    Best-effort loader for ProjectWise PowerShell cmdlets (pwps / pwps_dab).
    .DESCRIPTION
    Dashboard/watcher runs with -NoProfile on some hosts (servers, services), so modules that
    used to be present via user profile are missing. Try to import modules explicitly and
    provide actionable diagnostics if unavailable. pwps provides connection cmdlets while
    pwps_dab provides folder/document discovery; both may need to be re-imported in a worker
    runspace after opening a connection.
    #>
    [CmdletBinding()]
    param()

    $hasConnection = [bool](Get-Command -Name Open-PWConnection -ErrorAction SilentlyContinue)
    $hasDiscovery = _PWC-TestPWDiscoveryCmdlets
    if ($hasConnection -and $hasDiscovery) { return $true }

    foreach ($name in @('pwps_dab','pwps')) {
        try { Import-Module $name -Force -ErrorAction Stop | Out-Null } catch { }
        $hasConnection = [bool](Get-Command -Name Open-PWConnection -ErrorAction SilentlyContinue)
        $hasDiscovery = _PWC-TestPWDiscoveryCmdlets
        if ($hasConnection -and $hasDiscovery) { return $true }
    }

    # Common explicit path (Bentley install) if modules aren't in PSModulePath.
    $pwpsPath = 'C:\Program Files (x86)\Bentley\ProjectWise\bin\PowerShell\pwps\pwps.psd1'
    if (Test-Path -LiteralPath $pwpsPath) {
        try { Import-Module $pwpsPath -Force -ErrorAction Stop | Out-Null } catch { }
    }

    return [bool](Get-Command -Name Open-PWConnection -ErrorAction SilentlyContinue)
}

function Test-PWDiscoveryCmdlets {
    <#
    .SYNOPSIS
    Returns true when pwps_dab discovery/listing cmdlets are available in this runspace.
    #>
    [CmdletBinding()]
    param()

    return (_PWC-TestPWDiscoveryCmdlets)
}

function Ensure-PWDiscoveryCmdlets {
    <#
    .SYNOPSIS
    Ensures ProjectWise discovery/listing cmdlets from pwps_dab are available.
    .DESCRIPTION
    Open-PWConnection can succeed with only pwps loaded, or the runspace can effectively lose
    pwps_dab after connect. Force re-import pwps_dab and fail distinctly if folder/document
    discovery is still missing so callers do not mistake cmdlet loss for an empty datasource.
    #>
    [CmdletBinding()]
    param()

    if (_PWC-TestPWDiscoveryCmdlets) {
        return New-QCSuccessResult -Code 'PW_DISCOVERY_READY' -Message 'ProjectWise discovery cmdlets are available.' -Data @{ discoveryCmdlets = @('Get-PWFolderView','Get-PWDocumentsBySearch') }
    }

    try { Import-Module pwps_dab -Force -ErrorAction Stop | Out-Null } catch { }

    if (_PWC-TestPWDiscoveryCmdlets) {
        return New-QCSuccessResult -Code 'PW_DISCOVERY_READY' -Message 'ProjectWise discovery cmdlets are available after re-import.' -Data @{ discoveryCmdlets = @('Get-PWFolderView','Get-PWDocumentsBySearch') }
    }

    $msg = 'ProjectWise connected, but pwps_dab discovery cmdlets are missing (Get-PWFolderView/Get-PWDocumentsBySearch). Re-import pwps_dab or run the worker under ProjectWise PowerShell/MTA; do not treat this as an empty folder.'
    return New-QCFailureResult -Code 'PW_DISCOVERY_INCOMPLETE' -Message $msg -Data @{ psModulePath = $env:PSModulePath; missingDiscoveryCmdlets = @('Get-PWFolderView','Get-PWDocumentsBySearch') }
}

function _PWC-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Get-PWCredentialFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CredentialPath
    )

    if (-not (Test-Path -LiteralPath $CredentialPath)) {
        return New-QCFailureResult -Code 'PW_CRED_MISSING_FILE' -Message "Credential file not found: $CredentialPath" -Data @{ path = $CredentialPath }
    }
    try {
        $lines = Get-Content -LiteralPath $CredentialPath -ErrorAction Stop
        $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
        $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
        if (-not $uLine -or -not $pLine) {
            return New-QCFailureResult -Code 'PW_CRED_BAD_FORMAT' -Message "Invalid format in credential file: $CredentialPath" -Data @{ path = $CredentialPath }
        }
        $user = ($uLine -split '=', 2)[1].Trim()
        $pass = ($pLine -split '=', 2)[1].Trim()
        if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
            return New-QCFailureResult -Code 'PW_CRED_BAD_FORMAT' -Message "Credential file missing username/password values: $CredentialPath" -Data @{ path = $CredentialPath }
        }
        $sec = ConvertTo-SecureString $pass -AsPlainText -Force
        $cred = [pscredential]::new($user, $sec)
        return New-QCSuccessResult -Code 'PW_CRED_LOADED' -Message 'Credential loaded.' -Data @{ credential = $cred; userName = $user }
    } catch {
        return New-QCFailureResult -Code 'PW_CRED_READ_FAILED' -Message 'Failed to read credential file.' -Data @{ path = $CredentialPath; errorMessage = $_.Exception.Message }
    }
}

function Connect-PW {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatasourceName,
        [Parameter(Mandatory)]
        [pscredential]$Credential
    )

    if (-not (_PWC-TryImportPWModules)) {
        $psmp = $env:PSModulePath
        $msg = 'ProjectWise cmdlets not loaded (Open-PWConnection missing). Import pwps_dab/pwps or run under ProjectWise PowerShell. If pwps_dab is installed but still missing, ensure the watcher/worker runs in MTA (pwps_dab requires -MTA).'
        return New-QCFailureResult -Code 'PW_MISSING_MODULE' -Message $msg -Data @{ psModulePath = $psmp }
    }

    try {
        Open-PWConnection -DatasourceName $DatasourceName -UserName $Credential.UserName -Password $Credential.Password -WarningAction SilentlyContinue | Out-Null
        $discRes = Ensure-PWDiscoveryCmdlets
        if (-not $discRes.IsSuccess) {
            return New-QCFailureResult -Code 'PW_DISCOVERY_INCOMPLETE' -Message $discRes.Message -Data @{ datasourceName = $DatasourceName; userName = $Credential.UserName; discovery = $discRes.Data }
        }
        return New-QCSuccessResult -Code 'PW_CONNECTED' -Message 'ProjectWise connected.' -Data @{ datasourceName = $DatasourceName; userName = $Credential.UserName; discovery = $discRes.Data }
    } catch {
        return New-QCFailureResult -Code 'PW_CONNECT_FAILED' -Message 'ProjectWise connection failed.' -Data @{ datasourceName = $DatasourceName; userName = $Credential.UserName; errorMessage = $_.Exception.Message }
    }
}

function Disconnect-PW {
    [CmdletBinding()]
    param()

    if (-not (_PWC-TryImportPWModules)) {
        $psmp = $env:PSModulePath
        $msg = 'ProjectWise cmdlets not loaded (Close-PWConnection missing). Import pwps_dab/pwps or run under ProjectWise PowerShell. If pwps_dab is installed but still missing, ensure the watcher/worker runs in MTA (pwps_dab requires -MTA).'
        return New-QCFailureResult -Code 'PW_MISSING_MODULE' -Message $msg -Data @{ psModulePath = $psmp }
    }

    try {
        Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
        return New-QCSuccessResult -Code 'PW_DISCONNECTED' -Message 'ProjectWise disconnected.' -Data @{}
    } catch {
        return New-QCFailureResult -Code 'PW_DISCONNECT_FAILED' -Message 'ProjectWise disconnect failed.' -Data @{ errorMessage = $_.Exception.Message }
    }
}

function Get-PWImmediateChildFolders {
    <#
    .SYNOPSIS
    Returns immediate child folder objects under a ProjectWise folder (read-only).

    .DESCRIPTION
    pwps_dab varies by path: Get-PWFoldersImmediateChildren often returns nothing under CADD\Sheets while
    Get-PWFolderView lists discipline subfolders. This aligns oneLevelDeep expansion with Sheets discovery.

    FolderPath MUST use the pw cmdlet convention: without a leading Documents\ segment.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if (-not (_PWC-TryImportPWModules)) {
        return @()
    }

    function _PwcGetFolderView([string]$p) {
        $cmd = Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Parameters.ContainsKey('InputFolder')) {
            $f = Get-PWFolders -FolderPath $p -JustOne -ErrorAction Stop
            if (-not $f) { throw "Folder not found: $p" }
            return ($f | Get-PWFolderView -ErrorAction Stop)
        }
        return (Get-PWFolderView -FolderPath $p -ErrorAction Stop)
    }

    function _PwcPwProp([object]$Obj, [string]$Name) {
        try {
            if ($null -eq $Obj -or -not $Obj.PSObject -or -not $Obj.PSObject.Properties[$Name]) { return $null }
            return $Obj.$Name
        } catch { return $null }
    }

    $normalized = (($FolderPath -as [string]).Trim()).TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $collected = [System.Collections.Generic.List[object]]::new()

    $view = $null
    try {
        $view = _PwcGetFolderView $normalized
    } catch {
        try {
            $folderObj = Get-PWFolders -FolderPath $normalized -JustOne -ErrorAction SilentlyContinue
            if ($folderObj) {
                try {
                    $view = _PwcGetFolderView $normalized
                } catch { }
            }
        } catch { }
    }

    if ($view -and $view.Folders) {
        foreach ($f in @($view.Folders)) {
            if ($f) { [void]$collected.Add($f) }
        }
    }

    if ($collected.Count -eq 0 -and $view -and $view.Children) {
        foreach ($c in @($view.Children)) {
            if (-not $c) { continue }
            $docId = _PwcPwProp $c 'DocumentID'
            if (-not ([string]::IsNullOrWhiteSpace([string]$docId))) {
                continue
            }
            if (-not (_PwcPwProp $c 'FolderPath')) {
                continue
            }
            [void]$collected.Add($c)
        }
    }

    if ($collected.Count -gt 0) {
        return @($collected.ToArray())
    }

    try {
        return @(Get-PWFoldersImmediateChildren -FolderPath $normalized -WarningAction SilentlyContinue -ErrorAction Stop)
    } catch {
        try {
            return @(Get-PWFoldersImmediateChildren -FolderPath $normalized -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        } catch {
            return @()
        }
    }
}

# PW.Connection.psm1
# Responsibility: ProjectWise connection health and reconnect logic wrappers.


function Invoke-PWAuthenticatedCommand {
    <#
    .SYNOPSIS
    Runs a script block inside a ProjectWise session opened from a key/value credential file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatasourceName,
        [Parameter(Mandatory)]
        [string]$CredentialPath,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )

    $credRes = Get-PWCredentialFromFile -CredentialPath $CredentialPath
    if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }

    $conn = Connect-PW -DatasourceName $DatasourceName -Credential ([pscredential]$credRes.Data.credential)
    if (-not $conn.IsSuccess) { throw ($conn.Code + ': ' + $conn.Message) }

    try {
        return (& $ScriptBlock)
    } finally {
        [void](Disconnect-PW)
    }
}

function Get-PWFolderViewCompat {
    <#
    .SYNOPSIS
    Returns a ProjectWise folder view while handling pwps_dab parameter differences.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if (-not (_PWC-TryImportPWModules)) {
        throw 'ProjectWise cmdlets not loaded (Get-PWFolderView missing). Import pwps_dab/pwps or run under ProjectWise PowerShell.'
    }

    $folderViewCmd = Get-Command -Name Get-PWFolderView -ErrorAction Stop
    if ($folderViewCmd.Parameters.ContainsKey('InputFolder')) {
        $f = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction Stop
        if (-not $f) { throw "Folder not found: $FolderPath" }
        return ($f | Get-PWFolderView -ErrorAction Stop)
    }

    return (Get-PWFolderView -FolderPath $FolderPath -ErrorAction Stop)
}

function Get-PWFolderViewChildren {
    <#
    .SYNOPSIS
    Splits a ProjectWise folder view into folder and document collections.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$FolderView
    )

    $folders = @()
    if ($FolderView.Folders) { $folders = @($FolderView.Folders) }
    elseif ($FolderView.Children) { $folders = @($FolderView.Children | Where-Object { $_ -and $_.FolderPath -and -not $_.DocumentID }) }

    $docs = @()
    if ($FolderView.Documents) { $docs = @($FolderView.Documents) }
    elseif ($FolderView.Children) { $docs = @($FolderView.Children | Where-Object { $_ -and ($_.DocumentID -or $_.Name) }) }

    return [pscustomobject]@{
        Folders   = $folders
        Documents = $docs
    }
}

function Get-PWDocumentsInFolderRaw {
    <#
    .SYNOPSIS
    Returns documents directly inside a ProjectWise folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if (-not (_PWC-TryImportPWModules)) {
        throw 'ProjectWise cmdlets not loaded (Get-PWDocumentsBySearch missing). Import pwps_dab/pwps or run under ProjectWise PowerShell.'
    }

    return @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
}

function Show-PWFolderBrowser {
    <#
    .SYNOPSIS
    Prints child folders and documents for a ProjectWise folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatasourceName,
        [Parameter(Mandatory)]
        [string]$CredentialPath,
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory = $false)]
        [int]$Max = 40
    )

    Invoke-PWAuthenticatedCommand -DatasourceName $DatasourceName -CredentialPath $CredentialPath -ScriptBlock {
        Write-Host ("FolderPath: {0}" -f $FolderPath) -ForegroundColor Cyan
        $view = Get-PWFolderViewCompat -FolderPath $FolderPath
        $children = Get-PWFolderViewChildren -FolderView $view

        Write-Host ("Folders: {0}" -f @($children.Folders).Count) -ForegroundColor Green
        $children.Folders | Select-Object -First $Max Name,FolderPath | Format-Table -AutoSize
        Write-Host ("Documents: {0}" -f @($children.Documents).Count) -ForegroundColor Green
        $children.Documents | Select-Object -First $Max Name,Description,DocumentID,FullPath | Format-Table -AutoSize
    }
}

function Show-PWFolderDocumentList {
    <#
    .SYNOPSIS
    Prints documents and PDF counts for a ProjectWise folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatasourceName,
        [Parameter(Mandatory)]
        [string]$CredentialPath,
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory = $false)]
        [int]$Max = 25
    )

    Invoke-PWAuthenticatedCommand -DatasourceName $DatasourceName -CredentialPath $CredentialPath -ScriptBlock {
        Write-Host ("FolderPath: {0}" -f $FolderPath) -ForegroundColor Cyan
        $docs = @(Get-PWDocumentsInFolderRaw -FolderPath $FolderPath)
        Write-Host ("TotalDocs: {0}" -f $docs.Count) -ForegroundColor Green

        $pdfs = @($docs | Where-Object { $_.Name -match '\.pdf$' })
        Write-Host ("PdfDocs:   {0}" -f $pdfs.Count) -ForegroundColor Green

        $pdfs | Select-Object -First $Max Name,DocumentID,Description,FullPath | Format-Table -AutoSize
    }
}

function _PWC-ConvertProbeFolderPath {
    param([AllowNull()][string]$InternalFolderPath)

    $s = ($InternalFolderPath -as [string]).Trim().TrimEnd('\')
    while ($s -match '^(?i)Documents\\') { $s = $s -replace '^(?i)Documents\\', '' }
    if ([string]::IsNullOrWhiteSpace($s)) { return 'Documents' }
    return $s
}

function _PWC-GetSessionProbeFolderCandidates {
    param(
        [hashtable]$Config,
        [string]$ProbeFolderPath = ''
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    [void]$candidates.Add('Documents')

    if (-not [string]::IsNullOrWhiteSpace($ProbeFolderPath)) {
        $converted = _PWC-ConvertProbeFolderPath -InternalFolderPath $ProbeFolderPath
        if ($candidates -notcontains $converted) { [void]$candidates.Add($converted) }
    }

    $pwCfg = $Config.projectWise
    if ($pwCfg -is [pscustomobject]) {
        $tmp = @{}
        foreach ($p in $pwCfg.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
        $pwCfg = $tmp
    }
    if ($pwCfg -and $pwCfg.watchList) {
        $watchList = $pwCfg.watchList
        if ($watchList -is [pscustomobject]) {
            $wl = @{}
            foreach ($p in $watchList.PSObject.Properties) { $wl[$p.Name] = $p.Value }
            $watchList = $wl
        }
        if ($watchList.roots) {
            foreach ($root in @($watchList.roots)) {
                $path = $null
                if ($root -is [hashtable] -and $root.path) { $path = [string]$root.path }
                elseif ($root.path) { $path = [string]$root.path }
                if ([string]::IsNullOrWhiteSpace($path)) { continue }
                $converted = _PWC-ConvertProbeFolderPath -InternalFolderPath $path
                if ($candidates -notcontains $converted) { [void]$candidates.Add($converted) }
            }
        }
    }

    return @($candidates)
}

function Test-PWLoginHealth {
    <#
    .SYNOPSIS
    Checks ProjectWise login/session health.
    .DESCRIPTION
    Verifies connection/session viability for read-only operations.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only connectivity checks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$ProbeFolderPath = ''
    )

    if (-not (_PWC-TryImportPWModules)) {
        return New-QCFailureResult -Code 'PW_SESSION_UNHEALTHY' -Message 'ProjectWise cmdlets not loaded for session health probe.' -Data @{
            psModulePath = $env:PSModulePath
        }
    }

    if (Get-Command -Name 'Ensure-PWDiscoveryCmdlets' -ErrorAction SilentlyContinue) {
        $discRes = Ensure-PWDiscoveryCmdlets
        if (-not $discRes.IsSuccess) {
            return New-QCFailureResult -Code 'PW_SESSION_UNHEALTHY' -Message $discRes.Message -Data @{
                discoveryCode = [string]$discRes.Code
                discovery = $discRes.Data
            }
        }
    }

    $probeCandidates = @(_PWC-GetSessionProbeFolderCandidates -Config $Config -ProbeFolderPath $ProbeFolderPath)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $attempts = @()
    $lastError = ''

    foreach ($probePath in $probeCandidates) {
        try {
            $folder = Get-PWFolders -FolderPath $probePath -JustOne -ErrorAction Stop
            if ($folder) {
                $sw.Stop()
                return New-QCSuccessResult -Code 'PW_SESSION_HEALTHY' -Message 'ProjectWise session probe succeeded.' -Data @{
                    probeFolderPath = $probePath
                    probeCandidates = $probeCandidates
                    durationMs = [int]$sw.ElapsedMilliseconds
                    folderName = if ($folder.Name) { [string]$folder.Name } else { '' }
                }
            }
            $attempts += @{ probeFolderPath = $probePath; result = 'empty' }
            $lastError = "Get-PWFolders returned no folder for '$probePath'."
        } catch {
            $attempts += @{ probeFolderPath = $probePath; result = 'error'; errorMessage = [string]$_.Exception.Message }
            $lastError = [string]$_.Exception.Message
        }
    }

    $sw.Stop()
    return New-QCFailureResult -Code 'PW_SESSION_UNHEALTHY' -Message 'ProjectWise session probe failed after trying configured folder paths.' -Data @{
        probeFolderPath = if ($probeCandidates.Count -gt 0) { [string]$probeCandidates[0] } else { 'Documents' }
        probeCandidates = $probeCandidates
        attempts = $attempts
        durationMs = [int]$sw.ElapsedMilliseconds
        errorMessage = $lastError
    }
}

function Connect-PWIfNeeded {
    <#
    .SYNOPSIS
    Ensures a valid ProjectWise session exists.
    .DESCRIPTION
    Establishes or refreshes connection only as needed for subsequent read operations.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: session management calls; no document writes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

Export-ModuleMember -Function @(
    # Public API (used by watcher/worker/status-set)
    'Get-PWCredentialFromFile',
    'Connect-PW',
    'Disconnect-PW',
    'Test-PWDiscoveryCmdlets',
    'Ensure-PWDiscoveryCmdlets',
    'Get-PWImmediateChildFolders',
    'Invoke-PWAuthenticatedCommand',
    'Get-PWFolderViewCompat',
    'Get-PWFolderViewChildren',
    'Get-PWDocumentsInFolderRaw',
    'Show-PWFolderBrowser',
    'Show-PWFolderDocumentList',
    'Test-PWLoginHealth',
    'Connect-PWIfNeeded'
)
