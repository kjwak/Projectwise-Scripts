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
$sec = ConvertTo-SecureString $map['password'] -AsPlainText -Force
$ds = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
Open-PWConnection -DatasourceName $ds -UserName $map['username'] -Password $sec -Admin -WarningAction SilentlyContinue | Out-Null

try {
  $usr = Get-PWUser -UserName 'SVC_TYPSA_Archivist'
  Write-Output '=== User.Settings object ==='
  if ($usr.Settings) {
    Write-Output ("Settings type: {0}" -f $usr.Settings.GetType().FullName)
    $usr.Settings | Format-List * | Out-String | Write-Output
    Write-Output 'Settings property names:'
    $usr.Settings.PSObject.Properties.Name -join ', ' | Write-Output
    foreach ($p in @($usr.Settings.PSObject.Properties)) {
      $n = $p.Name
      if ($n -match '(?i)cred|expir|pass|auth|session|login|dialog|sso|ims|token|timeout|policy') {
        Write-Output ("  {0} = {1}" -f $n, $p.Value)
      }
    }
  }

  Write-Output ''
  Write-Output '=== Get-PWUserSettingByUser CredentialExpirationPolicy ==='
  $settingNames = @(
    'General_CredentialExpirationPolicy',
    'General_CanChangeGeneralSettings',
    'UI_ShowDialogOnError',
    'General_CanOnlyLoginThroughWebViewServer'
  )
  foreach ($sn in $settingNames) {
    try {
      $r = Get-PWUserSettingByUser -InputUsers $usr -SettingName $sn -ErrorAction Stop
      Write-Output ("{0}:" -f $sn)
      $r | Format-List * | Out-String | Write-Output
    } catch {
      Write-Output ("{0}: ERROR {1}" -f $sn, $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Get-PWUserSetting ByUserID with categories ==='
  # Discover UserSettingCategory enum/values if possible
  $cmd = Get-Command Get-PWUserSetting
  $p = $cmd.Parameters['UserSettingCategory']
  if ($p) {
    Write-Output ("Category param type: {0}" -f $p.ParameterType.FullName)
    if ($p.ParameterType.IsEnum) {
      [enum]::GetNames($p.ParameterType) | ForEach-Object { Write-Output ("  enum: {0}" -f $_) }
    }
  }
  foreach ($cat in @('General','UI','Admin','NetworkTransfer','Doc')) {
    try {
      $vals = Get-PWUserSetting -UserID 790 -UserSettingCategory $cat -ErrorAction Stop
      Write-Output ("--- category {0} ---" -f $cat)
      $vals | Where-Object {
        $_.UserSettingName -match '(?i)cred|expir|pass|auth|session|login|dialog|sso|policy' -or
        $_.Name -match '(?i)cred|expir|pass|auth|session|login|dialog|sso|policy'
      } | Format-List * | Out-String | Write-Output
      # also print CredentialExpiration if present in full list
      $vals | Where-Object { ($_ | Out-String) -match 'Credential' } | Format-List * | Out-String | Write-Output
    } catch {
      Write-Output ("category {0}: {1}" -f $cat, $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Compare defaults vs user for CredentialExpirationPolicy ==='
  try {
    $def = Get-PWUserDefaultSettings -SettingName 'General_CredentialExpirationPolicy' -ErrorAction Stop
    Write-Output 'DEFAULT:'
    $def | Format-List * | Out-String | Write-Output
  } catch { Write-Output ("default err: {0}" -f $_.Exception.Message) }

  Write-Output ''
  Write-Output '=== Datasource notes ==='
  Write-Output 'typsa-us-pw-03 IsSSO=True IsSTS=True'
  Write-Output 'User Type=Logical, Identity empty, no Windows identity, no federated identity'
  Write-Output 'Connect path used by automation: Open-PWConnection NoGUI UserName/Password'

  Write-Output ''
  Write-Output '=== Get-PWUserIdentityProvider (system) ==='
  Get-PWUserIdentityProvider | Format-List * | Out-String | Write-Output
}
finally {
  Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
