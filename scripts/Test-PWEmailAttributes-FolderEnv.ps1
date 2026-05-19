param([string]$FolderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1')
$ErrorActionPreference = 'Stop'
$lines = Get-Content 'C:\PW_QC_LOCAL\pw_cred.txt'
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred = [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps, pwps_dab -Force
Open-PWConnection -DatasourceName 'typsa-us-pw.bentley.com:typsa-us-pw-03' -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
try {
    $folder = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction Stop
    Write-Host '=== Folder properties (env-related) ==='
    $folder.PSObject.Properties | Where-Object { $_.Name -match 'Env|Workflow|Attribute' } | ForEach-Object {
        Write-Host "$($_.Name) = $($_.Value)"
    }


    $pdf = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First 1
    Write-Host "`n=== Document env-related props ==="
    $pdf.PSObject.Properties | Where-Object { $_.Name -match 'Env|Attribute|Workflow|State' } | ForEach-Object {
        Write-Host "$($_.Name) = $($_.Value)"
    }

    Write-Host "`n=== Get-PWDocumentsBySearchExtended ==="
    if (Get-Command Get-PWDocumentsBySearchExtended -ErrorAction SilentlyContinue) {
        $ext = Get-PWDocumentsBySearchExtended -FolderPath $FolderPath -JustThisFolder -DocumentName $pdf.Name -ErrorAction Stop | Select-Object -First 1
        $ext.PSObject.Properties | Where-Object { $_.Name -match 'EM_|Env|Attribute|Designer|Reviewer' } | ForEach-Object {
            Write-Host "$($_.Name) = $($_.Value)"
        }
    }

    Write-Host "`n=== CEL probe (if available) ==="
    if (Get-Command Invoke-PWCelExpression -ErrorAction SilentlyContinue) {
        foreach ($expr in @(
            'EM_Designer_Email',
            'EM_Reviewer_Email',
            '@this.EM_Designer_Email',
            '@this.EM_Reviewer_Email'
        )) {
            try {
                $r = Invoke-PWCelExpression -Expression $expr -ContextVersion $pdf -ErrorAction Stop
                Write-Host "CEL [$expr] = $r"
            } catch {
                try {
                    $r = Invoke-PWCelExpression -Expression $expr -ErrorAction Stop
                    Write-Host "CEL [$expr] = $r"
                } catch {
                    Write-Host "CEL [$expr] error: $($_.Exception.Message)"
                }
            }
        }
    } else {
        Write-Host 'Invoke-PWCelExpression not available'
    }

    Write-Host "`n=== ReturnColumns with Caltrans column set ==="
    $caltransCols = @('EM_Designer_Email', 'EM_Reviewer_Email', 'PE_NO', 'FIRM', 'a_attrno')
    $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $pdf.Name -ColumnsToReturn $caltransCols -ErrorAction Stop | Select-Object -First 1
    foreach ($cn in $caltransCols) {
        if ($w.PSObject.Properties[$cn]) { Write-Host "$cn = [$($w.$cn)]" } else { Write-Host "$cn = (property missing)" }
    }
    if ($w.Attributes) {
        foreach ($bag in @($w.Attributes)) {
            if ($bag -is [System.Collections.IDictionary] -and $bag.Count -gt 0) {
                foreach ($k in $bag.Keys) { Write-Host "Attributes[$k]=$($bag[$k])" }
            }
        }
    }
} finally {
    Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
