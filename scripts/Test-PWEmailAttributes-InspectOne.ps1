# Quick read-only inspect of one document's e-attributes in Seg_1
param(
    [string]$FolderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1',
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
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
    $pick = @($docs | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First 1)
    if (-not $pick) { $pick = $docs[0] }
    Write-Host "DOC: $($pick.Name) DocumentID=$($pick.DocumentID) ProjectID=$($pick.ProjectID)"
    $ea = @(Get-PWDocumentEAttributes -DocumentID $pick.DocumentID -ProjectID $pick.ProjectID)
    Write-Host "EATTR COUNT: $($ea.Count)"
    $i = 0
    foreach ($a in $ea) {
        $i++
        Write-Host "--- eattr #$i type=$($a.GetType().FullName) ---"
        if ($a -is [System.Collections.IDictionary]) {
            foreach ($k in $a.Keys) { Write-Host "  [$k] = $($a[$k])" }
        }
        foreach ($p in $a.PSObject.Properties) {
            Write-Host "  $($p.Name) = $($p.Value)"
        }
    }
    $cols = @('Name','DocumentID','EM_Designer_Email','EM_Reveiewer_Email','EM_Reviewer_Email')
    $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $pick.Name -ColumnsToReturn $cols -ErrorAction Stop | Select-Object -First 1
    Write-Host '--- search return columns ---'
    foreach ($cn in $cols) {
        if ($w.PSObject.Properties[$cn]) { Write-Host "  $cn = $($w.$cn)" }
    }
    Write-Host '--- all properties containing EM or mail ---'
    $w.PSObject.Properties | Where-Object { $_.Name -match 'EM_|mail|Designer|Reviewer|Reveiew' } | ForEach-Object {
        Write-Host "  $($_.Name) = $($_.Value)"
    }
} finally {
    Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
