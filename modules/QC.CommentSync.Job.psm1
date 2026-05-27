# QC.CommentSync.Job.psm1
# Responsibility: Build/enrich QC_COMMENT_STATUS_SYNC job metadata payloads.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.PdfExport.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PW.Connection.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue

function _QCJ-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function New-QCCommentSyncJobMetadata {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [hashtable]$Candidate = $null,
        [string]$LocalDownloadPath = '',
        [string]$SourcePwState = '',
        [string]$ContentFileHash = ''
    )

    if (-not $Candidate -and $Job.metadata -and $Job.metadata.candidate) {
        $Candidate = $Job.metadata.candidate
    }
    if (-not ($Candidate -is [hashtable])) { $Candidate = @{} }

    $fileHash = $ContentFileHash
    if (_QCJ-IsNullOrWhiteSpace $fileHash -and $Candidate.file -and $Candidate.file.sha256) {
        $fileHash = [string]$Candidate.file.sha256
    }

    $modified = ''
    if ($Candidate.file -and $Candidate.file.lastWriteTimeUtc) {
        $modified = [string]$Candidate.file.lastWriteTimeUtc
    }

    $docGuid = ''
    if ($Candidate.documentGuid) { $docGuid = [string]$Candidate.documentGuid }

    $meta = @{
        documentId = $docGuid
        documentGuid = $docGuid
        pwPath = [string]$Job.sourcePath
        localDownloadPath = $LocalDownloadPath
        fileName = [string]$Job.sourceName
        sourceModifiedUtc = $modified
        fileHash = $fileHash
        triggerRuleId = if ($Job.triggerRule -and $Job.triggerRule.id) { [string]$Job.triggerRule.id } else { '' }
        projectId = if ($Candidate.watchRoot) { [string]$Candidate.watchRoot } elseif ($Candidate.datasourceName) { [string]$Candidate.datasourceName } else { '' }
        watchRoot = if ($Candidate.watchRoot) { [string]$Candidate.watchRoot } else { '' }
        sourcePwState = $SourcePwState
        triggerSource = if ($Candidate.triggerSource) { [string]$Candidate.triggerSource } else { '' }
    }

    if (-not $Job.ContainsKey('metadata') -or -not ($Job.metadata -is [hashtable])) {
        $Job['metadata'] = @{}
    }
    $Job.metadata['commentSync'] = $meta
    return $meta
}

function Get-QCCommentSyncJobMetadata {
    param([hashtable]$Job)
    if ($Job.metadata -and $Job.metadata.commentSync) { return $Job.metadata.commentSync }
    return @{}
}

function Get-QCCommentSyncPwDocument {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    if (-not (Get-Command -Name 'Invoke-PWAuthenticatedCommand' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_COMMENT_SYNC_PW_UNAVAILABLE' -Message 'PW.Connection not available.' -Data @{}
    }

    $folder = [string]$Job.sourceFolder
    $name = [string]$Job.sourceName
    if (_QCJ-IsNullOrWhiteSpace $folder) {
        try { $folder = [System.IO.Path]::GetDirectoryName([string]$Job.sourcePath) } catch { }
    }
    if (_QCJ-IsNullOrWhiteSpace $name) { $name = [System.IO.Path]::GetFileName([string]$Job.sourcePath) }

    $ds = $null
    if ($Config.projectWise -and $Config.projectWise.datasourceName) { $ds = [string]$Config.projectWise.datasourceName }
    $credPath = $null
    if ($Config.projectWise -and $Config.projectWise.credentialPath) { $credPath = [string]$Config.projectWise.credentialPath }

    $meta = Get-QCCommentSyncJobMetadata -Job $Job
    $guid = if ($meta.documentGuid) { [string]$meta.documentGuid } else { '' }

    $block = {
        $doc = $null
        if (-not [string]::IsNullOrWhiteSpace($script:qcDocGuid)) {
            # Prefer GUID lookup to avoid path normalization issues.
            $cmdByGuids = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
            if ($cmdByGuids) {
                try {
                    $byGuid = @(Get-PWDocumentsByGUIDs -DocumentGUIDs @($script:qcDocGuid) -ErrorAction SilentlyContinue)
                    if ($byGuid -and $byGuid.Count -gt 0) { $doc = $byGuid[0] }
                } catch { }
            }
        }

        if (-not $doc -and (Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue)) {
            # Normalize internal folder path ("Documents\...") to what PW cmdlets accept.
            $apiFolder = $null
            if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
                try { $apiFolder = ConvertTo-PWCmdletFolderPath -InternalFolderPath $script:qcFolder } catch { }
            }
            if ([string]::IsNullOrWhiteSpace($apiFolder)) { $apiFolder = $script:qcFolder }

            foreach ($p in @($apiFolder, ('Documents\' + $apiFolder.TrimStart('\')))) {
                if ([string]::IsNullOrWhiteSpace($p)) { continue }
                try {
                    $found = Get-PWDocumentsBySearch -FolderPath $p -JustThisFolder -DocumentName $script:qcDocName -PopulatePath -ErrorAction Stop
                    if ($found) { $doc = @($found) | Select-Object -First 1; break }
                } catch { }
            }
        }

        if (-not $doc -and (Get-Command -Name 'Get-PWDocumentsInFolder' -ErrorAction SilentlyContinue)) {
            $folderForList = $script:qcFolder
            if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
                try {
                    $tmp = ConvertTo-PWCmdletFolderPath -InternalFolderPath $script:qcFolder
                    if (-not [string]::IsNullOrWhiteSpace($tmp)) { $folderForList = $tmp }
                } catch { }
            }
            $docs = @(Get-PWDocumentsInFolder -FolderPath $folderForList)
            foreach ($d in $docs) {
                $n = $null
                if (Get-Command Get-PWDocName -ErrorAction SilentlyContinue) { $n = Get-PWDocName -Doc $d }
                elseif ($d.Name) { $n = [string]$d.Name }
                if ($n -eq $script:qcDocName) { $doc = $d; break }
            }
        }
        if (-not $doc) {
            return New-QCFailureResult -Code 'QC_COMMENT_SYNC_DOC_NOT_FOUND' -Message 'PW document not found for comment sync.' -Data @{
                folder = $script:qcFolder; name = $script:qcDocName
            }
        }
        return New-QCSuccessResult -Code 'QC_COMMENT_SYNC_DOC_OK' -Message 'PW document resolved.' -Data @{ document = $doc }
    }

    try {
        $script:qcFolder = $folder
        $script:qcDocName = $name
        $script:qcDocGuid = $guid
        return Invoke-PWAuthenticatedCommand -DatasourceName $ds -CredentialPath $credPath -ScriptBlock $block
    } catch {
        return New-QCFailureResult -Code 'QC_COMMENT_SYNC_PW_FAILED' -Message $_.Exception.Message -Data @{}
    } finally {
        Remove-Variable -Name qcFolder, qcDocName, qcDocGuid -Scope Script -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function New-QCCommentSyncJobMetadata, Get-QCCommentSyncJobMetadata, Get-QCCommentSyncPwDocument
