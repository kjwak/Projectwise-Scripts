Get-Item "C:\Program Files\Bentley\ProjectWise\bin\dmscli.dll" |
Select-Object FullName,
@{n="FileVersion";e={$_.VersionInfo.FileVersion}},
@{n="ProductVersion";e={$_.VersionInfo.ProductVersion}}

Get-Item "C:\Program Files\Bentley\ProjectWise\bin\aaApi.dll" |
Select-Object FullName,
@{n="FileVersion";e={$_.VersionInfo.FileVersion}},
@{n="ProductVersion";e={$_.VersionInfo.ProductVersion}}

Get-PWCurrentDatasource -ErrorAction SilentlyContinue