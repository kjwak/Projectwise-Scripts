$ErrorActionPreference='Stop'
$lines=Get-Content 'C:\PW_QC_LOCAL\pw_cred.txt'
$user=(($lines|Where-Object{$_ -match 'username'}|Select -First 1)-split '=',2)[1].Trim()
$pass=(($lines|Where-Object{$_ -match 'password'}|Select -First 1)-split '=',2)[1].Trim()
$cred=[pscredential]::new($user,(ConvertTo-SecureString $pass -AsPlainText -Force))
Import-Module pwps,pwps_dab -Force
Open-PWConnection -DatasourceName 'typsa-us-pw.bentley.com:typsa-us-pw-03' -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue|Out-Null
try {
  $raw = Get-PWEnvironments -ErrorAction Stop
  $envs = @()
  if ($raw -is [System.Collections.IEnumerable] -and -not ($raw -is [string])) {
    foreach ($e in $raw) { if ($e -and $e.PSObject.Properties['Name']) { $envs += $e } }
  }
  if ($envs.Count -eq 0) { $envs = @($raw) }
  Write-Host "Environments: $($envs.Count)"
  foreach ($env in $envs) {
    $name = if ($env.Name) { $env.Name } else { $env.EnvironmentName }
    Write-Host "ENV: $name"
    try {
      $cols = @(Get-PWEnvironmentColumns -EnvironmentName $name -ErrorAction Stop)
      Write-Host "  columns: $($cols.Count)"
      $cols | Where-Object {
        $cn = if ($_.Name) { $_.Name } else { $_.ColumnName }
        $cn -match 'EM_|Email|Designer|Reviewer|Reveiew'
      } | ForEach-Object {
        $cn = if ($_.Name) { $_.Name } else { $_.ColumnName }
        Write-Host "    $cn"
      }
    } catch { Write-Host "  columns error: $($_.Exception.Message)" }
  }
} finally { Close-PWConnection -EA SilentlyContinue|Out-Null}
