$ErrorActionPreference = 'Stop'
Import-Module pwps_dab -ErrorAction Stop
$dsName = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
$credPath = 'C:\PW_QC_LOCAL\pw_cred.txt'
$lines = Get-Content -LiteralPath $credPath
$user = (($lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = ConvertTo-SecureString ((($lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1) -split '=', 2)[1].Trim()) -AsPlainText -Force
Open-PWConnection -DatasourceName $dsName -UserName $user -Password $pass | Out-Null

$folderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Ref'
$docs = @(Get-PWDocumentsBySearch -FolderPath $folderPath -JustThisFolder)
$dgns = @($docs | Where-Object { [string]$_.Name -like '*.dgn' -or [string]$_.FileName -like '*.dgn' })
$sample = $dgns | Where-Object { $_.Name -eq '0818000063M-ROW-TYA.dgn' } | Select-Object -First 1

$r0 = @(Get-PWReferenceFiles -InputDocument $sample -ReferenceLevel 0)
$rAll = @(Get-PWReferenceFiles -InputDocument $sample -ReferenceLevel -1)
Write-Host ("LEVEL0=" + $r0.Count)
$r0 | ForEach-Object { Write-Host ("L0 " + $_.Name) }
Write-Host ("ALL=" + $rAll.Count)
$rAll | ForEach-Object { Write-Host ("ALL " + $_.Name) }

$rpt = Get-PWReferenceReportForDocuments -InputDocument $sample
Write-Host ("REPORT_ROWS=" + $rpt.Rows.Count)

# Second sample that likely has more nested refs
$sample2 = $dgns | Where-Object { $_.Name -like '*DRN*' } | Select-Object -First 1
Write-Host ("SAMPLE2=" + $sample2.Name)
$r02 = @(Get-PWReferenceFiles -InputDocument $sample2 -ReferenceLevel 0)
$rAll2 = @(Get-PWReferenceFiles -InputDocument $sample2 -ReferenceLevel -1)
Write-Host ("S2_LEVEL0=" + $r02.Count + " S2_ALL=" + $rAll2.Count)
$r02 | ForEach-Object { Write-Host ("S2L0 " + $_.Name) }
$rAll2 | ForEach-Object { Write-Host ("S2ALL " + $_.Name) }

Write-Host 'DONE'
