# Verify EM_Designer_Email / EM_Reviewer_Email on Caltrans environment docs in Seg_1
param([string]$FolderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1')
$ErrorActionPreference = 'Stop'
$cols = @('EM_Designer_Email', 'EM_Reviewer_Email')
$lines = Get-Content 'C:\PW_QC_LOCAL\pw_cred.txt'
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred = [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps, pwps_dab -Force
Open-PWConnection -DatasourceName 'typsa-us-pw.bentley.com:typsa-us-pw-03' -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
try {
    $pdfs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath | Where-Object { $_.Name -match '\.pdf$' })
    $hits = 0
    $samples = @()
    foreach ($d in $pdfs) {
        $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $d.Name -ColumnsToReturn $cols -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $w) { continue }
        $designer = if ($w.PSObject.Properties['EM_Designer_Email']) { [string]$w.EM_Designer_Email } else { '' }
        $reviewer = if ($w.PSObject.Properties['EM_Reviewer_Email']) { [string]$w.EM_Reviewer_Email } else { '' }
        if ($designer -or $reviewer) {
            $hits++
            if ($samples.Count -lt 8) {
                $samples += [pscustomobject]@{ Name = $d.Name; EM_Designer_Email = $designer; EM_Reviewer_Email = $reviewer }
            }
        }
    }
    Write-Host "PDFs with EM email values: $hits / $($pdfs.Count)"
    $samples | Format-Table -AutoSize

    $d0 = $pdfs[0]
    $w0 = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $d0.Name -ColumnsToReturn $cols | Select-Object -First 1
    Write-Host "First PDF ($($d0.Name)):"
    Write-Host "  EM_Designer_Email=[$($w0.EM_Designer_Email)]"
    Write-Host "  EM_Reviewer_Email=[$($w0.EM_Reviewer_Email)]"
    Write-Host "  Property exists: Designer=$($w0.PSObject.Properties['EM_Designer_Email'] -ne $null) Reviewer=$($w0.PSObject.Properties['EM_Reviewer_Email'] -ne $null)"

    # EAttributes expanded
    $raw = Get-PWDocumentEAttributes -DocumentID $d0.DocumentID -ProjectID $d0.ProjectID -ErrorAction SilentlyContinue
    if ($raw -is [System.Collections.IEnumerable]) {
        foreach ($item in $raw) {
            if ($item -is [System.Collections.IDictionary]) {
                foreach ($k in $item.Keys) {
                    if ([string]$k -match 'EM_') { Write-Host "  eattr $k=$($item[$k])" }
                }
            } elseif ($item.PSObject.Properties['Name'] -and [string]$item.Name -match 'EM_') {
                Write-Host "  eattr $($item.Name)=$($item.Value)"
            }
        }
    }
} finally {
    Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
