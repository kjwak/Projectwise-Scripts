# test_qc_rendition.ps1 — profile resolution, readiness gate, notification deferral.

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Processing\QC.Rendition.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Notifications\QC.Notifications.psm1') -Force

function Assert-Eq($Actual, $Expected, [string]$Label) {
    if ("$Actual" -cne "$Expected") {
        throw "$Label expected '$Expected' but got '$Actual'"
    }
}

function Assert-True($Value, [string]$Label) {
    if (-not $Value) { throw $Label }
}

$readinessRoot = Join-Path $env:TEMP ('qc_rendition_test_' + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $readinessRoot -Force | Out-Null

try {
    $config = @{
        projectWise = @{
            datasourceName = 'ds-primary'
            watchList = @{
                roots = @(
                    @{
                        path = 'Documents\ProjectX'
                        qcRendition = @{
                            profileName = 'Profile A'
                            outputFolderRelative = 'Renditions'
                        }
                    }
                )
            }
        }
        qcWorkflow = @{
            enabled = $true
            states = @{ qcReceived = 'Ready for QC'; readyForQc = 'Ready for QC' }
        }
        qcRendition = @{
            enabled = $true
            deferReadyForQcNotification = $true
            readinessStorePath = $readinessRoot
            folderOverrides = @(
                @{
                    folderPathPrefix = 'ProjectX\CADD\Sheets\Special'
                    profileName = 'Profile B'
                    outputFolderRelative = 'QC PDF'
                }
            )
        }
        notifications = @{
            enabled = $true
            provider = 'Mock'
            dryRun = $true
            events = @{
                'Ready for QC' = @{
                    enabled = $true
                    eventType = 'READY_FOR_QC'
                    to = @('reviewers')
                    subjectTemplate = 'Ready for QC - {documentName}'
                }
            }
        }
    }

    $res = Resolve-QCRenditionProfile -Config $config -DatasourceName 'ds-primary' -FolderPath 'ProjectX\Other\CADD\Sheets'
    Assert-True $res.IsSuccess 'Resolve watch root profile'
    Assert-Eq $res.Data.profile.profileName 'Profile A' 'Watch root profile'

    $res2 = Resolve-QCRenditionProfile -Config $config -DatasourceName 'ds-primary' -FolderPath 'Documents\ProjectX\CADD\Sheets\Special'
    Assert-Eq $res2.Data.profile.profileName 'Profile B' 'Folder override profile'
    Assert-Eq $res2.Data.profile.outputFolderRelative 'QC PDF' 'Folder override output'

    $sameFolderProfile = @{ outputFolderRelative = '.' }
    $outSame = _QCR-ResolveOutputFolderPath -Profile $sameFolderProfile -ApiFolderPath 'AZDOT\Proj\CADD\Sheets'
    Assert-Eq $outSame 'AZDOT\Proj\CADD\Sheets' 'Dot means same sheet folder'

    $key = Get-QCReadinessKey -DocumentGuid 'abc-123' -FolderPath 'f' -QcPdfName 'sheet-qc.pdf'
    Assert-Eq $key 'guid:abc-123' 'Guid readiness key'

    Set-QCReadinessFlag -Config $config -ReadinessKey $key -PrependComplete | Out-Null
    $state = Get-QCReadinessState -Config $config -ReadinessKey $key
    Assert-True $state.prependComplete 'Prepend flagged'
    Assert-True (-not $state.renditionComplete) 'Rendition not yet'

    Assert-True (Test-QCShouldDeferReadyForQcNotification -Config $config -CurrentState 'Ready for QC') 'Defer Ready for QC notification'
    Assert-True (-not (Test-QCShouldDeferReadyForQcNotification -Config $config -CurrentState 'Redlines Received')) 'No defer other states'

    $defer = Invoke-QCNotificationForStateChange -Config $config -PreviousState 'In Production' -CurrentState 'Ready for QC' `
        -DocumentName 'sheet-qc.pdf' -DocumentGuid 'abc-123'
    Assert-Eq $defer.Code 'QC_NOTIFICATION_DEFERRED_READY_FOR_QC' 'Notification deferred'

    Set-QCReadinessFlag -Config $config -ReadinessKey $key -RenditionComplete | Out-Null
    $sent = Invoke-QCReadyForQcNotificationIfReady -Config $config -ReadinessKey $key `
        -PreviousState 'In Production' -DocumentName 'sheet-qc.pdf' -DocumentGuid 'abc-123'
    Assert-True (($sent.IsSuccess) -or ($sent.Code -eq 'QC_NOTIFICATION_SKIPPED_NO_RECIPIENTS')) 'Ready notification path invoked'

    $parentJob = @{
        id = 'prepend-job-1'
        type = 'QC_PREPEND'
        sourcePath = 'C:\x\sheet-qc.pdf'
        sourceName = 'sheet-qc.pdf'
        sourceFolder = 'ProjectX\CADD\Sheets'
        dedupeKey = 'dq_test'
        triggerRule = @{ id = 'r1'; jobType = 'QC_PREPEND' }
    }
    $built = New-QCRenditionQueueJob -Config $config -ParentJob $parentJob -DocumentGuid 'abc-123'
    Assert-Eq $built.Data.job.type 'QC_RENDITION' 'Built rendition job type'
    Assert-Eq $built.Data.job.metadata.rendition.parentPrependJobId 'prepend-job-1' 'Parent job link'

    $sheetKey = Get-QCRenditionSheetReadinessKey -FolderPath 'ProjectX\CADD\Sheets' -SourceDgnFileName 'sheet.dgn'
    Assert-Eq $sheetKey 'sheet:projectx\cadd\sheets|sheet' 'Sheet readiness key'

    $dk1 = _QCR-GetRenditionDedupeKeyForSheet -FolderPath 'ProjectX\CADD\Sheets' -SourceDgnFileName 'sheet.dgn'
    $dk2 = _QCR-GetRenditionDedupeKeyForSheet -FolderPath 'ProjectX\CADD\Sheets' -SourceDgnFileName 'sheet.dgn'
    Assert-Eq $dk1 $dk2 'Sheet dedupe key stable'
    Assert-True ($dk1.StartsWith('dq_qcrendition_')) 'Sheet dedupe prefix'

    $jid1 = _QCR-GetRenditionJobIdForSheet -FolderPath 'ProjectX\CADD\Sheets' -SourceDgnFileName 'sheet.dgn'
    $jid2 = _QCR-GetRenditionJobIdForSheet -FolderPath 'ProjectX\CADD\Sheets' -SourceDgnFileName 'sheet.dgn'
    Assert-Eq $jid1 $jid2 'Sheet job id stable'

    Write-Host 'All QC rendition tests passed.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $readinessRoot) {
        Remove-Item -LiteralPath $readinessRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
