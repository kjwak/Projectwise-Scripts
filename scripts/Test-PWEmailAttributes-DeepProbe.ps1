# Deep read-only probe for EM email fields on Seg_1 documents
param(
    [string]$FolderPath = 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1',
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
)
$ErrorActionPreference = 'Continue'

function Expand-EAttributes([object]$Raw) {
    if ($null -eq $Raw) { return @() }
    $items = [System.Collections.Generic.List[object]]::new()
    if ($Raw -is [System.Collections.IEnumerable] -and -not ($Raw -is [string])) {
        foreach ($x in $Raw) {
            if ($null -eq $x) { continue }
            if ($x.GetType().FullName -like '*EAttribute*') { [void]$items.Add($x) }
            elseif ($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])) {
                foreach ($y in (Expand-EAttributes $x)) { [void]$items.Add($y) }
            }
        }
    }
    return @($items.ToArray())
}

function Get-EmProps([object]$Obj) {
    $out = @{}
    if (-not $Obj) { return $out }
    foreach ($p in $Obj.PSObject.Properties) {
        if ($p.Name -match 'EM_|Designer|Reviewer|Reveiew|mail') {
            $out[$p.Name] = $p.Value
        }
    }
    return $out
}

$lines = Get-Content -LiteralPath $CredentialPath
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred = [pscredential]::new($user, (ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps, pwps_dab -Force
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
try {
    $docs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath)
    Write-Host "Total docs: $($docs.Count)"

    # Document object properties (search result)
    Write-Host "`n=== Search object EM_* props (first 5 docs) ==="
    foreach ($d in ($docs | Select-Object -First 5)) {
        $em = Get-EmProps $d
        Write-Host "$($d.Name): $(if ($em.Count) { ($em.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ' } else { '(none)' })"
    }

    # EAttributes on PDFs and DGNs - count non-empty
    foreach ($ext in @('pdf', 'dgn')) {
        $subset = @($docs | Where-Object { $_.Name -match "\.$ext`$" })
        $nonEmpty = 0
        $samples = @()
        foreach ($doc in $subset) {
            $raw = Get-PWDocumentEAttributes -DocumentID $doc.DocumentID -ProjectID $doc.ProjectID -ErrorAction SilentlyContinue
            $attrs = Expand-EAttributes $raw
            if ($attrs.Count -gt 0) {
                $nonEmpty++
                if ($samples.Count -lt 3) {
                    $pairs = @{}
                    foreach ($a in $attrs) {
                        if ($a.PSObject.Properties['Name']) { $pairs[[string]$a.Name] = [string]$a.Value }
                    }
                    $samples += [pscustomobject]@{ Name = $doc.Name; AttrCount = $attrs.Count; Pairs = ($pairs | ConvertTo-Json -Compress) }
                }
            }
        }
        Write-Host "`n=== .$ext eattr non-empty: $nonEmpty / $($subset.Count) ==="
        $samples | Format-List
    }

    # Extended search if available
    if (Get-Command Get-PWDocumentsBySearchExtended -ErrorAction SilentlyContinue) {
        Write-Host "`n=== Get-PWDocumentsBySearchExtended (first PDF) ==="
        $pdf = @($docs | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First 1)
        if ($pdf) {
            try {
                $ext = Get-PWDocumentsBySearchExtended -FolderPath $FolderPath -JustThisFolder -DocumentName $pdf.Name -ErrorAction Stop | Select-Object -First 1
                $em = Get-EmProps $ext
                if ($em.Count) { $em.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key)=$($_.Value)" } }
                else { Write-Host '  (no EM/mail props)' }
            } catch { Write-Host "  error: $($_.Exception.Message)" }
        }
    }

    # Try return columns with wildcard-ish broader set
    $pdf0 = @($docs | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First 1)
    if ($pdf0) {
        Write-Host "`n=== ReturnColumns all props for $($pdf0.Name) ==="
        $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $pdf0.Name -ColumnsToReturn @('*') -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $w) {
            $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $pdf0.Name -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        if ($w) {
            $w.PSObject.Properties | Sort-Object Name | ForEach-Object { Write-Host "  $($_.Name) = $($_.Value)" }
        } else { Write-Host '  no row returned' }
    }

    # Environment on document
    Write-Host "`n=== Environment-related props on first PDF ==="
    if ($pdf0) {
        $pdf0.PSObject.Properties | Where-Object { $_.Name -match 'Env|Environment' } | ForEach-Object {
            Write-Host "  $($_.Name) = $($_.Value)"
        }
    }
} finally {
    Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
