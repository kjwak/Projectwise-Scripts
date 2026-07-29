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
Open-PWConnection -DatasourceName 'typsa-us-pw.bentley.com:typsa-us-pw-03' -UserName $map['username'] -Password $sec -Admin -WarningAction SilentlyContinue | Out-Null

try {
  Write-Output '=== All settings containing Service/Credential/Login/Expire ==='
  $cats = [enum]::GetNames([Bentley.ProjectWise.PowerShell.Common.UserSetting+Category])
  foreach ($cat in $cats) {
    try {
      $vals = Get-PWUserSetting -UserID 790 -UserSettingCategory $cat -ErrorAction Stop
      foreach ($v in $vals) {
        if ($v.UserSettingName -match '(?i)service|cred|expir|login|token|license|delegate') {
          Write-Output ("{0}.{1} = {2}" -f $cat, $v.UserSettingName, $v.UserSettingValue)
        }
      }
    } catch {
      Write-Output ("{0}: {1}" -f $cat, $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Defaults for service-related settings ==='
  $def = Get-PWUserDefaultSettings
  $def | Where-Object { $_.UserSettingName -match '(?i)service|cred|expir|login|token|license|delegate' } |
    Format-Table UserSettingName, UserSettingValue -AutoSize | Out-String | Write-Output

  # What does CredentialExpirationPolicy value 0 / -1 / 'No expiration' look like on another user?
  Write-Output '=== Sample CredentialExpirationPolicy across users (first matches) ==='
  $users = Get-PWUser | Select-Object -First 200
  $found = 0
  foreach ($u in $users) {
    try {
      $pol = Get-PWUserSetting -UserID $u.ID -UserSettingCategory General -ErrorAction SilentlyContinue |
        Where-Object { $_.UserSettingName -eq 'CredentialExpirationPolicy' }
      if ($pol) {
        $val = [string]$pol.UserSettingValue
        if ($val -ne 'Default' -and $val -ne '0') {
          Write-Output ("{0} (id={1}, type={2}): CredentialExpirationPolicy={3}" -f $u.UserName, $u.ID, $u.Type, $val)
          $found++
          if ($found -ge 15) { break }
        }
      }
    } catch {}
  }
}
finally {
  Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
