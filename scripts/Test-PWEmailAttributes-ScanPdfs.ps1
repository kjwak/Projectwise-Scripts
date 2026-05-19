# Scan all PDFs in folder for EM_* email attributes (read-only)
param(
    [string]$FolderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1',
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
)
$ErrorActionPreference = 'Stop'
$TargetAttrs = @('EM_Designer_Email', 'EM_Reveiewer_Email', 'EM_Reviewer_Email')

function Expand-EAttributes([object]$Raw) {
    if ($null -eq $Raw) { return @() }
    if ($Raw -is [Bentley.ProjectWise.PowerShell.Common.EAttribute]) { return @($Raw) }
    $items = @()
    if ($Raw -is [System.Collections.IEnumerable] -and -not ($Raw -is [string])) {
        foreach ($x in $Raw) {
            if ($x -is [Bentley.ProjectWise.PowerShell.Common.EAttribute]) { $items += $x }
            elseif ($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])) {
                $items += Expand-EAttributes $x
            }
        }
    }
    return $items
}

function Read-EAttrPairs([object]$Raw) {
    $pairs = @{}
    foreach ($a in (Expand-EAttributes $Raw)) {
        $n = $null; $v = $null
        if ($a.PSObject.Properties['Name']) { $n = [string]$a.Name }
        if ($a.PSObject.Properties['Value']) { $v = [string]$a.Value }
        if ($n) { $pairs[$n] = $v }
    }
    return $pairs
}

$lines = Get-Content -LiteralPath $CredentialPath
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred = [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps, pwps_dab -Force
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
try {
    $pdfs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath | Where-Object { $_.Name -match '\.pdf$' })
    Write-Host "PDF count: $($pdfs.Count)"

    $hits = @()
    foreach ($doc in $pdfs) {
        $raw = Get-PWDocumentEAttributes -DocumentID $doc.DocumentID -ProjectID $doc.ProjectID -ErrorAction SilentlyContinue
        $pairs = Read-EAttrPairs $raw
        $match = @{}
        foreach ($k in $TargetAttrs) {
            if ($pairs.ContainsKey($k) -and -not [string]::IsNullOrWhiteSpace([string]$pairs[$k])) {
                $match[$k] = $pairs[$k]
            }
        }
        if ($match.Count -gt 0) {
            $hits += [pscustomobject]@{ Name = $doc.Name; DocumentID = $doc.DocumentID; Attributes = ($match | ConvertTo-Json -Compress) }
        }
    }
    Write-Host "PDFs with target attrs via Get-PWDocumentEAttributes: $($hits.Count)"
    $hits | Select-Object -First 10 | Format-Table -AutoSize

    # Try return columns on first 3 PDFs - dump ALL non-empty props
    Write-Host "`nReturnColumns probe (first 3 PDFs):"
    foreach ($doc in ($pdfs | Select-Object -First 3)) {
        $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $doc.Name -ColumnsToReturn $TargetAttrs -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $w) { Write-Host "  $($doc.Name): no row"; continue }
        $found = @($w.PSObject.Properties | Where-Object { $_.Name -in $TargetAttrs -and $_.Value } | ForEach-Object { "$($_.Name)=$($_.Value)" })
        Write-Host "  $($doc.Name): $($found -join ', ')"
        if (-not $found) {
            $em = @($w.PSObject.Properties | Where-Object { $_.Name -match 'EM_|mail|Designer|Reviewer|Reveiew' -and $_.Value })
            if ($em) { $em | ForEach-Object { Write-Host "    alt $($_.Name)=$($_.Value)" } }
            else { Write-Host '    (no EM/mail properties on object)' }
        }
    }

    # List environments and search for column definitions (bounded)
    if (Get-Command Get-PWEnvironments -ErrorAction SilentlyContinue) {
        Write-Host "`nEnvironment column search:"
        $envs = @(Get-PWEnvironments -ErrorAction SilentlyContinue | Select-Object -First 15)
        foreach ($env in $envs) {
            $envName = if ($env.Name) { $env.Name } elseif ($env.EnvironmentName) { $env.EnvironmentName } else { [string]$env }
            if ([string]::IsNullOrWhiteSpace($envName)) { continue }
            try {
                $cols = @(Get-PWEnvironmentColumns -EnvironmentName $envName -ErrorAction Stop)
            } catch {
                try { $cols = @(Get-PWEnvironmentColumns -Environment $envName -ErrorAction Stop) } catch { continue }
            }
            $m = @($cols | Where-Object {
                $n = if ($_.Name) { [string]$_.Name } else { [string]$_.ColumnName }
                $n -in $TargetAttrs -or $n -match 'Designer_Email|Reveiewer_Email|Reviewer_Email'
            })
            if ($m.Count -gt 0) {
                Write-Host "  Environment '$envName':"
                $m | ForEach-Object {
                    $n = if ($_.Name) { $_.Name } else { $_.ColumnName }
                    Write-Host "    $n"
                }
            }
        }
    }
} finally {
    Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
