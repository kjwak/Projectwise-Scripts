$ErrorActionPreference = 'Stop'
Import-Module pwps -Force
Import-Module pwps_dab -Force

$credPath = 'C:\PW_QC_LOCAL\pw_cred.txt'
$map = @{}
Get-Content -LiteralPath $credPath | ForEach-Object {
  $line = $_.Trim()
  if (-not $line -or $line.StartsWith('#')) { return }
  $i = $line.IndexOf('=')
  if ($i -lt 1) { return }
  $map[$line.Substring(0, $i).Trim().ToLowerInvariant()] = $line.Substring($i + 1).Trim()
}
$userName = $map['username']
$sec = ConvertTo-SecureString $map['password'] -AsPlainText -Force
$ds = 'typsa-us-pw.bentley.com:typsa-us-pw-03'

Write-Output '=== Connect with -Admin ==='
try {
  Open-PWConnection -DatasourceName $ds -UserName $userName -Password $sec -Admin -WarningAction SilentlyContinue | Out-Null
  Write-Output 'Admin connect OK'
} catch {
  Write-Output ("Admin connect FAILED: {0}" -f $_.Exception.Message)
  Write-Output 'Trying without -Admin for session cmdlets only...'
  Open-PWConnection -DatasourceName $ds -UserName $userName -Password $sec -WarningAction SilentlyContinue | Out-Null
}

try {
  Write-Output ''
  Write-Output '=== Get-PWLoginStatus / Get-PWCurrentDSSession / Get-PWSessions ==='
  foreach ($name in @('Get-PWLoginStatus','Get-PWCurrentDSSession','Get-PWSessions')) {
    Write-Output ("--- {0} ---" -f $name)
    try {
      $r = & $name -ErrorAction Stop
      $r | Format-List * | Out-String | Write-Output
    } catch {
      Write-Output ("  error: {0}" -f $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Get-PWCurrentUser (full) ==='
  $cu = Get-PWCurrentUser
  $cu | Format-List * | Out-String | Write-Output

  Write-Output ''
  Write-Output '=== Get-PWUser SVC_TYPSA_Archivist ==='
  try {
    $usr = Get-PWUser -UserName 'SVC_TYPSA_Archivist' -ErrorAction Stop
    Write-Output ("property count: {0}" -f @($usr.PSObject.Properties).Count)
    $usr | Format-List * | Out-String | Write-Output
  } catch {
    Write-Output ("Get-PWUser error: {0}" -f $_.Exception.Message)
  }

  Write-Output ''
  Write-Output '=== Get-PWUserIdentity for user ==='
  try {
    $ids = @(Get-PWUserIdentity -UserName 'SVC_TYPSA_Archivist' -ErrorAction Stop)
    Write-Output ("count={0}" -f $ids.Count)
    $ids | Format-List * | Out-String | Write-Output
  } catch {
    try {
      $ids = @(Get-PWUserIdentity -ErrorAction Stop)
      Write-Output ("all identities count={0}" -f $ids.Count)
      $ids | Where-Object { $_.UserName -match 'Archivist|TYPSA' -or $_.Name -match 'Archivist|TYPSA' } | Format-List * | Out-String | Write-Output
      if ($ids.Count -gt 0 -and -not ($ids | Where-Object { $_.UserName -match 'Archivist|TYPSA' -or $_.Name -match 'Archivist|TYPSA' })) {
        Write-Output 'Sample identity object properties:'
        $ids[0].PSObject.Properties.Name -join ', ' | Write-Output
        $ids | Select-Object -First 5 | Format-List * | Out-String | Write-Output
      }
    } catch {
      Write-Output ("Get-PWUserIdentity error: {0}" -f $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Get-PWUserWindowsIdentity for user ==='
  try {
    $w = @(Get-PWUserWindowsIdentity -UserName 'SVC_TYPSA_Archivist' -ErrorAction Stop)
    Write-Output ("count={0}" -f $w.Count)
    $w | Format-List * | Out-String | Write-Output
  } catch {
    Write-Output ("error: {0}" -f $_.Exception.Message)
    try {
      Get-Command Get-PWUserWindowsIdentity | ForEach-Object {
        $_.ParameterSets | ForEach-Object { Write-Output ("  set {0}: {1}" -f $_.Name, (($_.Parameters | ForEach-Object Name) -join ', ')) }
      }
    } catch {}
  }

  Write-Output ''
  Write-Output '=== Get-PWUserSetting / Get-PWUserDefaultSettings / Get-PWUserSettingByUser ==='
  foreach ($name in @('Get-PWUserSetting','Get-PWUserDefaultSettings','Get-PWUserSettingByUser')) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if (-not $c) { Write-Output ("(missing) {0}" -f $name); continue }
    Write-Output ("--- {0} parameter sets ---" -f $name)
    foreach ($ps in $c.ParameterSets) {
      Write-Output ("  [{0}] {1}" -f $ps.Name, (($ps.Parameters | ForEach-Object Name) -join ', '))
    }
    Write-Output ("--- {0} invoke ---" -f $name)
    try {
      if ($name -eq 'Get-PWUserSettingByUser') {
        & $name -UserName 'SVC_TYPSA_Archivist' -ErrorAction Stop | Format-List * | Out-String | Write-Output
      } elseif ($name -eq 'Get-PWUserSetting') {
        # try common patterns
        try { & $name -UserName 'SVC_TYPSA_Archivist' -ErrorAction Stop | Format-List * | Out-String | Write-Output }
        catch {
          try { & $name -ErrorAction Stop | Select-Object -First 50 | Format-List * | Out-String | Write-Output }
          catch { Write-Output ("  error: {0}" -f $_.Exception.Message) }
        }
      } else {
        & $name -ErrorAction Stop | Format-List * | Out-String | Write-Output
      }
    } catch {
      Write-Output ("  error: {0}" -f $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Target datasource SSO/STS flags ==='
  $dsObj = Get-PWDatasource | Where-Object { $_.FullName -eq $ds -or $_.InternalName -eq 'typsa-us-pw-03' }
  if ($dsObj) { $dsObj | Format-List * | Out-String | Write-Output } else { Write-Output 'target ds not in list (may still be connected)' }

  Write-Output ''
  Write-Output '=== Native identity cmds ==='
  foreach ($name in @('Get-PWUserNativeIdentity','Get-PWNonFederatedLoginToken')) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if (-not $c) { Write-Output ("(missing) {0}" -f $name); continue }
    Write-Output ("--- {0} ---" -f $name)
    foreach ($ps in $c.ParameterSets) {
      Write-Output ("  [{0}] {1}" -f $ps.Name, (($ps.Parameters | ForEach-Object Name) -join ', '))
    }
    try {
      & $name -UserName 'SVC_TYPSA_Archivist' -ErrorAction Stop | Format-List * | Out-String | Write-Output
    } catch {
      try { & $name -ErrorAction Stop | Format-List * | Out-String | Write-Output }
      catch { Write-Output ("  error: {0}" -f $_.Exception.Message) }
    }
  }
}
finally {
  try { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null } catch {}
  Write-Output 'Disconnected.'
}
