param([string]$FolderPath='Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1')
$ErrorActionPreference='Stop'
$lines=Get-Content 'C:\PW_QC_LOCAL\pw_cred.txt'
$user = (($lines | Where-Object { $_ -match 'username' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$pass = (($lines | Where-Object { $_ -match 'password' } | Select-Object -First 1) -split '=', 2)[1].Trim()
$cred=[pscredential]::new($user,(ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps,pwps_dab -Force
Open-PWConnection -DatasourceName 'typsa-us-pw.bentley.com:typsa-us-pw-03' -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue|Out-Null
try {
  $pdf = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath | Where-Object { $_.Name -match '\.pdf$' } | Select-Object -First 1
  $w=Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $pdf.Name|Select -First 1
  Write-Host "Doc: $($pdf.Name)"
  Write-Host "Attributes type: $($w.Attributes.GetType().FullName)"
  Write-Host "Attributes Count: $($w.Attributes.Count)"
  if ($w.Attributes -is [System.Collections.IEnumerable] -and -not ($w.Attributes -is [string])) {
    $idx = 0
    foreach ($bag in @($w.Attributes)) {
      $idx++
      Write-Host "Attributes bag #$idx type=$($bag.GetType().FullName) count=$($bag.Count)"
      if ($bag -is [System.Collections.IDictionary]) {
        foreach ($k in $bag.Keys) { Write-Host "  KEY=[$k] VAL=[$($bag[$k])]" }
      } else {
        $bag.PSObject.Properties | ForEach-Object { Write-Host "  $($_.Name)=$($_.Value)" }
      }
    }
  }
  Write-Host "CustomAttributes type: $($w.CustomAttributes.GetType().FullName)"
  Write-Host "CustomAttributes Count: $($w.CustomAttributes.Count)"
  if ($w.CustomAttributes -is [System.Collections.IDictionary]) {
    foreach ($k in $w.CustomAttributes.Keys) { Write-Host "CKEY=[$k] CVAL=[$($w.CustomAttributes[$k])]" }
  }
  # all envs for EM columns
  if (Get-Command Get-PWEnvironments -EA SilentlyContinue) {
    $found=@()
    foreach ($env in @(Get-PWEnvironments -EA SilentlyContinue)) {
      $en = if ($env.Name) { $env.Name } else { $env.EnvironmentName }
      if (-not $en) { continue }
      try { $cols=@(Get-PWEnvironmentColumns -EnvironmentName $en -EA Stop) } catch { continue }
      foreach ($c in $cols) {
        $cn = if ($c.Name) { $c.Name } else { $c.ColumnName }
        if ($cn -match 'EM_Designer|EM_Reveiew|EM_Reviewer|Designer_Email|Reveiewer|Reviewer_Email') { $found += "$en :: $cn" }
        if ($cn -match 'Email' -and $cn -match 'Designer|Reviewer|Reveiew') { $found += "$en :: $cn" }
      }
    }
    Write-Host "Matching env columns: $($found.Count)"
    $found | Select-Object -First 20 | ForEach-Object { Write-Host $_ }

    $emAll = @()
    foreach ($env in @(Get-PWEnvironments -ErrorAction SilentlyContinue)) {
      $en = if ($env.Name) { $env.Name } else { $env.EnvironmentName }
      if (-not $en) { continue }
      try { $cols = @(Get-PWEnvironmentColumns -EnvironmentName $en -ErrorAction Stop) } catch { continue }
      foreach ($c in $cols) {
        $cn = if ($c.Name) { [string]$c.Name } else { [string]$c.ColumnName }
        if ($cn -match '^EM_') { $emAll += "$en :: $cn" }
      }
    }
    Write-Host "EM_* env columns (first 30): $($emAll.Count) total"
    $emAll | Select-Object -First 30 | ForEach-Object { Write-Host $_ }
  }

  Write-Host '--- QC Received PDFs (workflow state probe) ---'
  $pdfs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath | Where-Object { $_.Name -match '\.pdf$' })
  foreach ($p in $pdfs) {
    $w = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -DocumentName $p.Name -ColumnsToReturn @('Name','Workflow','WorkflowState','StateName','WorkflowName') -ErrorAction SilentlyContinue | Select-Object -First 1
    $state = if ($w.WorkflowState) { $w.WorkflowState } elseif ($w.StateName) { $w.StateName } else { '' }
    if ($state -match 'QC Received') {
      Write-Host "QC Received doc: $($p.Name)"
      $raw = Get-PWDocumentEAttributes -DocumentID $p.DocumentID -ProjectID $p.ProjectID -ErrorAction SilentlyContinue
      if ($raw) {
        foreach ($bag in @($raw)) {
          if ($bag -is [System.Collections.IDictionary]) {
            foreach ($k in $bag.Keys) { Write-Host "  eattr $k=$($bag[$k])" }
          }
        }
      }
    }
  }
} finally { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null }
