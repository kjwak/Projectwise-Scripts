$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Processing\QC.StatusSetPublication.psm1') -Force
Import-Module (Join-Path $root 'modules\Core\Core.Hashing.psm1') -Force

$failures = 0
function Assert-True([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

$work = Join-Path $env:TEMP ('sspub_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null
$queue = Join-Path $work 'queue'
New-Item -ItemType Directory -Path (Join-Path $queue '_watcher') -Force | Out-Null

$nFolder = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\N_Seg'
$sFolder = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\S_Seg'
$destFolder = 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\Out\_I-15-ELPSE_Civil-StatusSets'
$otherFolder = 'Documents\AZDOT\OtherProject\CADD\Sheets\Main'
$otherDest = 'Documents\AZDOT\OtherProject\Out\StatusSets'
$asOf = [datetime]::ParseExact('2026-11-02', 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)

function New-PdfFile([string]$Name, [string]$Body) {
    $path = Join-Path $work $Name
    $text = "%PDF-1.4`n$Body`n%%EOF`n"
    [System.IO.File]::WriteAllText($path, $text)
    return $path
}

function New-Store {
    return @{
        docs = New-Object System.Collections.Generic.List[hashtable]
        creates = New-Object System.Collections.Generic.List[hashtable]
        failOnce = New-Object System.Collections.Generic.List[string]
    }
}

function Add-StoreDoc($Store, [string]$Folder, [string]$Name, [string]$FilePath, [datetime]$LastWriteUtc) {
    $key = ConvertTo-StatusSetPublicationFolderKey -Path $Folder
    [void]$Store.docs.Add(@{
        FolderKey = $key
        Name = $Name
        Path = $FilePath
        LastWriteUtc = $LastWriteUtc
    })
}

function New-Gateway($Store) {
    return @{
        FindDocument = {
            param($FolderPath, $DocumentName)
            $key = ConvertTo-StatusSetPublicationFolderKey -Path $FolderPath
            foreach ($d in @($script:ActiveStore.docs)) {
                if ([string]$d.FolderKey -eq $key -and [string]$d.Name -eq [string]$DocumentName) {
                    $len = [int64](Get-Item -LiteralPath ([string]$d.Path)).Length
                    return @{
                        Name = [string]$d.Name
                        FileSize = $len
                        FileUpdateDateUtc = [datetime]$d.LastWriteUtc
                        _Path = [string]$d.Path
                    }
                }
            }
            return $null
        }
        ExportDocument = {
            param($Document, $TargetFolder)
            $src = [string]$Document['_Path']
            $name = [string]$Document['Name']
            if (-not (Test-Path -LiteralPath $TargetFolder)) { New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null }
            $dest = Join-Path $TargetFolder $name
            Copy-Item -LiteralPath $src -Destination $dest -Force
            return $dest
        }
        CreateDocument = {
            param($FolderPath, $FilePath, $DocumentName)
            if ($script:ActiveStore.failOnce.Contains([string]$DocumentName)) {
                [void]$script:ActiveStore.failOnce.Remove([string]$DocumentName)
                throw "Simulated copy failure for $DocumentName"
            }
            $saved = Join-Path $script:StoreDir (([guid]::NewGuid().ToString('N')) + '_' + [string]$DocumentName)
            Copy-Item -LiteralPath $FilePath -Destination $saved -Force
            Add-StoreDoc -Store $script:ActiveStore -Folder $FolderPath -Name $DocumentName -FilePath $saved -LastWriteUtc ([datetime]::UtcNow)
            [void]$script:ActiveStore.creates.Add(@{ Folder = [string]$FolderPath; Name = [string]$DocumentName })
        }
    }
}

function New-PublicationConfig([bool]$Enabled = $false, [bool]$ProjectEnabled = $true, [hashtable[]]$ExtraProjects = @()) {
    $i15 = @{
        projectKey = 'CAFWY2200-I-15_ELPSE'
        enabled = $ProjectEnabled
        sets = @(
            @{
                id = 'N_Seg'
                sourceFolder = $nFolder
                identification = @{ method = 'statusSetPdfName' }
                destinationFolder = $destFolder
                fileNamePattern = '{yyyyMMdd}_N-Seg_StatusSet.pdf'
            },
            @{
                id = 'S_Seg'
                sourceFolder = $sFolder
                identification = @{ method = 'statusSetPdfName' }
                destinationFolder = $destFolder
                fileNamePattern = '{yyyyMMdd}_S-Seg_StatusSet.pdf'
            }
        )
    }
    $projects = @($i15) + @($ExtraProjects)
    return @{
        dryRun = $false
        statusSet = @{ statusSetPdfName = '_StatusSet.pdf'; localRoot = (Join-Path $work 'localsets') }
        queue = @{ rootDir = $queue }
        statusSetPublication = @{
            enabled = $Enabled
            timeZoneId = 'America/Phoenix'
            trigger = @{ dayOfWeek = 'Monday'; time = '06:00' }
            projects = $projects
        }
    }
}

function Reset-ActiveStore {
    $script:StoreDir = Join-Path $work ('store_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:StoreDir -Force | Out-Null
    $script:ActiveStore = New-Store
    return (New-Gateway -Store $script:ActiveStore)
}

try {
    $nFresh = New-PdfFile 'n-fresh.pdf' 'north-fresh'
    $sFresh = New-PdfFile 's-fresh.pdf' 'south-fresh'
    $nOld = New-PdfFile 'n-old.pdf' 'north-old'
    $sOld = New-PdfFile 's-old.pdf' 'south-old'
    $freshUtc = [datetime]::ParseExact('2026-11-02T14:00:00Z', 'yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()
    $oldUtc = [datetime]::ParseExact('2026-10-20T12:00:00Z', 'yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()

    Write-Host 'Test: both status sets publish and verify' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    $nHashBefore = Get-Sha256FileHex -Path $nFresh
    $sHashBefore = Get-Sha256FileHex -Path $sFresh
    $cfg = New-PublicationConfig
    $r = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True $r.IsSuccess 'publish succeeds'
    Assert-True ([string]$r.Data.overall -eq 'SUCCESS') 'overall SUCCESS'
    Assert-True (@($script:ActiveStore.creates).Count -eq 2) 'two destination creates'
    $names = @($script:ActiveStore.creates | ForEach-Object { [string]$_.Name } | Sort-Object)
    Assert-True ($names -contains '20261102_N-Seg_StatusSet.pdf') 'north archive name'
    Assert-True ($names -contains '20261102_S-Seg_StatusSet.pdf') 'south archive name'
    $nSeg = @($r.Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    $sSeg = @($r.Data.segments | Where-Object { $_.id -eq 'S_Seg' })[0]
    Assert-True ([string]$nSeg.copy -eq 'SUCCESS' -and [string]$nSeg.verification -eq 'PASSED') 'north copy verified'
    Assert-True ([string]$sSeg.copy -eq 'SUCCESS' -and [string]$sSeg.verification -eq 'PASSED') 'south copy verified'
    Assert-True ((Get-Sha256FileHex -Path $nFresh) -eq $nHashBefore) 'north source bytes unchanged'
    Assert-True ((Get-Sha256FileHex -Path $sFresh) -eq $sHashBefore) 'south source bytes unchanged'
    Assert-True (-not [bool]$nSeg.newlyGenerated) 'latest-existing is not marked newly generated'

    Write-Host 'Test: preview performs no writes' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    $r = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Preview -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True $r.IsSuccess 'preview succeeds'
    Assert-True (@($script:ActiveStore.creates).Count -eq 0) 'preview created nothing'
    Assert-True (@($r.Data.segments | Where-Object { [string]$_.copy -eq 'NOT_RUN' -and [bool]$_.canPublish }).Count -eq 2) 'both segments publishable and not copied'

    Write-Host 'Test: manual publication outside the Monday window' -ForegroundColor Cyan
    $settings = (Get-StatusSetPublicationSettings -Config $cfg).Data
    $wed = Test-StatusSetPublicationSlot -Settings $settings -SlotKey 'full_scan_schedule|2026-11-04|06:00'
    $evening = Test-StatusSetPublicationSlot -Settings $settings -SlotKey 'full_scan_schedule|2026-11-02|18:00'
    $monday = Test-StatusSetPublicationSlot -Settings $settings -SlotKey 'full_scan_schedule|2026-11-02|06:00'
    Assert-True (-not [bool]$wed.matches) 'Wednesday 06:00 is not the publication slot'
    Assert-True (-not [bool]$evening.matches) 'Monday 18:00 is not the publication slot'
    Assert-True ([bool]$monday.matches) 'Monday 06:00 matches'
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nOld -LastWriteUtc $oldUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sOld -LastWriteUtc $oldUtc
    $wedDate = [datetime]::ParseExact('2026-11-04', 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $r = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $wedDate -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True ($r.IsSuccess -and [string]$r.Data.overall -eq 'SUCCESS') 'manual Wednesday publish succeeds'
    Assert-True (@($script:ActiveStore.creates | Where-Object { $_.Name -like '20261104_*' }).Count -eq 2) 'Wednesday file names used'
    Assert-True (@($r.Data.segments | Where-Object { -not [bool]$_.newlyGenerated }).Count -eq 2) 'old PDFs stay not newly generated'

    Write-Host 'Test: missing and invalid source PDFs' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    $bad = Join-Path $work 'not-pdf.txt'
    [System.IO.File]::WriteAllText($bad, 'hello')
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $bad -LastWriteUtc $freshUtc
    $r = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True (-not $r.IsSuccess) 'invalid and missing sources fail'
    $nSeg = @($r.Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    $sSeg = @($r.Data.segments | Where-Object { $_.id -eq 'S_Seg' })[0]
    Assert-True ([string]$nSeg.sourceStatus -eq 'INVALID') 'north invalid'
    Assert-True ([string]$sSeg.sourceStatus -eq 'MISSING') 'south missing'
    Assert-True (@($script:ActiveStore.creates).Count -eq 0) 'nothing copied'

    Write-Host 'Test: identical destination is already published' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $destFolder -Name '20261102_N-Seg_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $destFolder -Name '20261102_S-Seg_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    $r = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True ($r.IsSuccess -and [string]$r.Data.overall -eq 'SUCCESS') 'identical destination succeeds'
    Assert-True (@($script:ActiveStore.creates).Count -eq 0) 'identical destination was not rewritten'
    Assert-True (@($r.Data.segments | Where-Object { [string]$_.copy -eq 'ALREADY_PUBLISHED' }).Count -eq 2) 'both already published'

    Write-Host 'Test: conflicting destination is not overwritten' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    $otherBody = New-PdfFile 'other-body.pdf' 'different-bytes'
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $destFolder -Name '20261102_N-Seg_StatusSet.pdf' -FilePath $otherBody -LastWriteUtc $freshUtc
    $r = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True (-not $r.IsSuccess) 'conflict fails the run'
    $nSeg = @($r.Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    Assert-True ([bool]$nSeg.conflict -and [string]$nSeg.copy -eq 'NOT_RUN') 'north conflict was not overwritten'
    $conflictPath = @($script:ActiveStore.docs | Where-Object { $_.Name -eq '20261102_N-Seg_StatusSet.pdf' })[0].Path
    Assert-True ((Get-Sha256FileHex -Path $conflictPath) -eq (Get-Sha256FileHex -Path $otherBody)) 'conflict bytes unchanged'
    $sSeg = @($r.Data.segments | Where-Object { $_.id -eq 'S_Seg' })[0]
    Assert-True ([string]$sSeg.copy -eq 'SUCCESS') 'south still publishes independently'

    Write-Host 'Test: partial copy failure then recovery' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    [void]$script:ActiveStore.failOnce.Add('20261102_S-Seg_StatusSet.pdf')
    $first = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True (-not $first.IsSuccess) 'first run fails'
    $nSeg = @($first.Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    $sSeg = @($first.Data.segments | Where-Object { $_.id -eq 'S_Seg' })[0]
    Assert-True ([string]$nSeg.copy -eq 'SUCCESS') 'north completed before south failed'
    Assert-True ([string]$sSeg.copy -eq 'FAILED') 'south copy failed'
    $createsAfterFirst = @($script:ActiveStore.creates).Count
    $second = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True ($second.IsSuccess -and [string]$second.Data.overall -eq 'SUCCESS') 'retry succeeds'
    $n2 = @($second.Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    $s2 = @($second.Data.segments | Where-Object { $_.id -eq 'S_Seg' })[0]
    Assert-True ([string]$n2.copy -eq 'ALREADY_PUBLISHED') 'north was not copied again'
    Assert-True ([string]$s2.copy -eq 'SUCCESS' -and [string]$s2.verification -eq 'PASSED') 'south recovered'
    Assert-True (@($script:ActiveStore.creates).Count -eq ($createsAfterFirst + 1)) 'only the failed segment was copied on retry'

    Write-Host 'Test: invalid project configuration' -ForegroundColor Cyan
    $missing = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'DOES-NOT-EXIST' -AsOfDate $asOf -Mode Preview -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True (-not $missing.IsSuccess -and [string]$missing.Code -eq 'STATUS_SET_PUBLISH_CONFIG') 'unknown project fails'
    $badCfg = New-PublicationConfig
    $badCfg.statusSetPublication.projects = @(
        @{
            projectKey = 'CAFWY2200-I-15_ELPSE'
            enabled = $true
            sets = @(@{ id = 'N_Seg'; sourceFolder = $nFolder; identification = @{ method = 'bogus' }; destinationFolder = ''; fileNamePattern = 'no-date.pdf' })
        }
    )
    $createsBeforeBad = @($script:ActiveStore.creates).Count
    $badRun = Invoke-StatusSetPublication -Config $badCfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True (-not $badRun.IsSuccess) 'invalid set fails'
    Assert-True (@($script:ActiveStore.creates).Count -eq $createsBeforeBad) 'invalid config did not copy'

    $none = Get-StatusSetPublicationSettings -Config @{ dryRun = $false }
    Assert-True ($none.IsSuccess -and @($none.Data.projects).Count -eq 0 -and -not [bool]$none.Data.enabled) 'missing configuration schedules no publication'

    Write-Host 'Test: repeated manual execution' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    $once = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    $twice = Invoke-StatusSetPublication -Config $cfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True ($once.IsSuccess -and $twice.IsSuccess) 'both runs succeed'
    Assert-True (@($script:ActiveStore.creates).Count -eq 2) 'second run did not create again'
    Assert-True (@($twice.Data.segments | Where-Object { [string]$_.copy -eq 'ALREADY_PUBLISHED' }).Count -eq 2) 'second run reports already published'

    Write-Host 'Test: other projects are left alone' -ForegroundColor Cyan
    $other = @{
        projectKey = 'AZDOT-OTHER'
        enabled = $true
        sets = @(@{
            id = 'Main'
            sourceFolder = $otherFolder
            identification = @{ method = 'statusSetPdfName' }
            destinationFolder = $otherDest
            fileNamePattern = '{yyyyMMdd}_Main_StatusSet.pdf'
        })
    }
    $iso = New-PublicationConfig -ExtraProjects @($other)
    $gw = Reset-ActiveStore
    $otherPdf = New-PdfFile 'other-source.pdf' 'other-project'
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $otherFolder -Name '_StatusSet.pdf' -FilePath $otherPdf -LastWriteUtc $freshUtc
    $r = Invoke-StatusSetPublication -Config $iso -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -UseLatestExisting -RunKind Manual -DocumentGateway $gw
    Assert-True $r.IsSuccess 'i-15 publish succeeds beside another configured project'
    $otherKey = ConvertTo-StatusSetPublicationFolderKey -Path $otherDest
    $touchedOther = @($script:ActiveStore.docs | Where-Object { [string]$_.FolderKey -eq $otherKey -and $_.Name -ne '_StatusSet.pdf' })
    Assert-True ($touchedOther.Count -eq 0) 'other project destination was not written'
    Assert-True (@($script:ActiveStore.creates | Where-Object { [string]$_.Name -like '*Main*' }).Count -eq 0) 'no create used the other project name'

    Write-Host 'Test: scheduled slot does not publish when disabled and does not touch schedule state' -ForegroundColor Cyan
    $scheduleCalls = 0
    function Set-QCFullScanScheduleSlotComplete { $script:scheduleCalls++; return $true }
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    $off = Invoke-StatusSetScheduledPublication -Config (New-PublicationConfig -Enabled $false) -QueueRoot $queue -CompletedSlotKey 'full_scan_schedule|2026-11-02|06:00' -DocumentGateway $gw
    Assert-True ($off.IsSuccess -and [string]$off.Code -eq 'STATUS_SET_PUBLISH_DISABLED') 'disabled schedule is a no-op'
    Assert-True (@($script:ActiveStore.creates).Count -eq 0) 'disabled schedule did not copy'
    Assert-True ($script:scheduleCalls -eq 0) 'publication did not mark a reconciliation slot complete'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path (Join-Path $queue '_watcher') 'full-scan-schedule-state.json'))) 'schedule state file was not created'

    Write-Host 'Test: enabled Monday slot publishes; other slots do not' -ForegroundColor Cyan
    $onCfg = New-PublicationConfig -Enabled $true
    $skip = Invoke-StatusSetScheduledPublication -Config $onCfg -QueueRoot $queue -CompletedSlotKey 'full_scan_schedule|2026-11-04|06:00' -DocumentGateway $gw
    Assert-True ([string]$skip.Data.overall -eq 'SKIPPED') 'Wednesday slot skipped'
    Assert-True (@($script:ActiveStore.creates).Count -eq 0) 'skipped slot did not copy'
    $due = Invoke-StatusSetScheduledPublication -Config $onCfg -QueueRoot $queue -CompletedSlotKey 'full_scan_schedule|2026-11-02|06:00' -DocumentGateway $gw
    Assert-True ($due.IsSuccess -and [string]$due.Data.overall -eq 'SUCCESS') 'Monday slot publishes'
    Assert-True ([bool]$due.Data.scheduleStateModified -eq $false) 'schedule state flag stays false'
    Assert-True ($script:scheduleCalls -eq 0) 'Monday publish still did not complete the scan slot'
    $mondaySeg = @($due.Data.projects[0].Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    Assert-True ([bool]$mondaySeg.newlyGenerated) 'cycle-fresh source is newly generated'

    Write-Host 'Test: manual reconcile hints do not mark an unchanged set as newly generated' -ForegroundColor Cyan
    $hintCfg = New-PublicationConfig
    $rec = Invoke-StatusSetPublicationSourceReconcile -Config $hintCfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -ReconcileCommand {
        param($Job, $ReconcileConfig)
        $folder = ''
        if ($Job -is [hashtable] -and $Job.ContainsKey('sourceFolder')) { $folder = [string]$Job['sourceFolder'] }
        elseif ($Job -and $Job.PSObject.Properties['sourceFolder']) { $folder = [string]$Job.sourceFolder }
        if ($folder -like '*N_Seg') {
            return [pscustomobject]@{ IsSuccess = $true; Code = 'STATUS_SET_OK'; Message = 'rebuilt'; Data = @{} }
        }
        return [pscustomobject]@{ IsSuccess = $true; Code = 'STATUS_SET_UP_TO_DATE'; Message = 'current'; Data = @{} }
    }
    Assert-True $rec.IsSuccess 'manual reconcile wrapper succeeds'
    Assert-True ([bool]$rec.Data.scheduleStateModified -eq $false) 'reconcile wrapper reports schedule untouched'
    Assert-True ($script:scheduleCalls -eq 0) 'reconcile wrapper did not complete a scan slot'
    $nHint = $rec.Data.hints['N_Seg']
    $sHint = $rec.Data.hints['S_Seg']
    Assert-True ([bool]$nHint.newlyGenerated) ("rebuilt set is newly generated ({0}: {1})" -f [string]$nHint.reconcileCode, [string]$nHint.reconcileMessage)
    Assert-True (-not [bool]$sHint.newlyGenerated) ("current set is not newly generated ({0})" -f [string]$sHint.reconcileCode)
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nOld -LastWriteUtc $oldUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sOld -LastWriteUtc $oldUtc
    $pub = Invoke-StatusSetPublication -Config $hintCfg -ProjectKey 'CAFWY2200-I-15_ELPSE' -AsOfDate $asOf -Mode Publish -RunKind Manual -ProvenanceHints $rec.Data.hints -CycleStartUtc ([datetime]::UtcNow) -DocumentGateway $gw
    Assert-True $pub.IsSuccess 'hint-based publish succeeds'
    $nSeg = @($pub.Data.segments | Where-Object { $_.id -eq 'N_Seg' })[0]
    $sSeg = @($pub.Data.segments | Where-Object { $_.id -eq 'S_Seg' })[0]
    Assert-True ([bool]$nSeg.newlyGenerated -and [string]$nSeg.provenance -eq 'refreshed') 'north keeps the rebuild hint'
    Assert-True (-not [bool]$sSeg.newlyGenerated -and [string]$sSeg.provenance -eq 'confirmed_current') 'south is confirmed current, not newly generated'

    Write-Host 'Test: in-flight status set defers without copying' -ForegroundColor Cyan
    $gw = Reset-ActiveStore
    Add-StoreDoc -Store $script:ActiveStore -Folder $nFolder -Name '_StatusSet.pdf' -FilePath $nFresh -LastWriteUtc $freshUtc
    Add-StoreDoc -Store $script:ActiveStore -Folder $sFolder -Name '_StatusSet.pdf' -FilePath $sFresh -LastWriteUtc $freshUtc
    $deferred = Invoke-StatusSetScheduledPublication -Config $onCfg -QueueRoot $queue -CompletedSlotKey 'full_scan_schedule|2026-11-02|06:00' -DocumentGateway $gw -InFlightSetIds @('N_Seg', 'S_Seg')
    Assert-True ([string]$deferred.Data.overall -eq 'DEFERRED') 'in-flight work is deferred'
    Assert-True (@($script:ActiveStore.creates).Count -eq 0) 'deferred run did not copy'
    Assert-True (Test-Path -LiteralPath (Join-Path (Join-Path $queue '_watcher') 'statusset-publication-pending.json')) 'pending publication recorded'
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failures -gt 0) {
    Write-Host "test_statusset_publication.ps1 failed: $failures" -ForegroundColor Red
    exit 1
}
Write-Host 'test_statusset_publication.ps1 passed' -ForegroundColor Green
exit 0
