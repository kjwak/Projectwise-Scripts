param(
    [string]$FolderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1',
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',
    [int]$MaxDocs = 20
)
$ErrorActionPreference = 'Stop'
$lines = Get-Content -LiteralPath $CredentialPath
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred = [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps, pwps_dab -Force
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
try {
    $docs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath)
    $targets = @($docs | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First $MaxDocs)
    foreach ($d in $targets) {
        $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $d.Name -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $w) { continue }
        $hits = @()
        if ($w.Attributes -and $w.Attributes.Count -gt 0) {
            foreach ($k in $w.Attributes.Keys) {
                if ([string]$k -match 'EM_|Designer|Reviewer|Reveiew|mail') { $hits += "Attr[$k]=$($w.Attributes[$k])" }
            }
        }
        if ($w.CustomAttributes -and $w.CustomAttributes.Count -gt 0) {
            foreach ($k in $w.CustomAttributes.Keys) {
                if ([string]$k -match 'EM_|Designer|Reviewer|Reveiew|mail') { $hits += "Custom[$k]=$($w.CustomAttributes[$k])" }
            }
        }
        if ($hits.Count -gt 0) {
            Write-Host "$($d.Name): $($hits -join '; ')"
        }
    }
    Write-Host '--- First PDF full Attributes / CustomAttributes keys ---'
    $w0 = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $targets[0].Name -ErrorAction SilentlyContinue | Select-Object -First 1
    Write-Host "AttributeRecordCount=$($w0.AttributeRecordCount) Attributes.Count=$($w0.Attributes.Count) CustomAttributes.Count=$($w0.CustomAttributes.Count)"
    if ($w0.Attributes.Count -gt 0) {
        $w0.Attributes.Keys | Sort-Object | ForEach-Object { Write-Host "  Attr $_ = $($w0.Attributes[$_])" }
    }
    if ($w0.CustomAttributes.Count -gt 0) {
        $w0.CustomAttributes.Keys | Sort-Object | ForEach-Object { Write-Host "  Custom $_ = $($w0.CustomAttributes[$_])" }
    }
} finally { Close-PWConnection -EA SilentlyContinue | Out-Null }
