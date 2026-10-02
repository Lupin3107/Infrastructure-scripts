Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
Import-Module Microsoft.PowerApps.Administration.PowerShell
Add-PowerAppsAccount

Get-AdminPowerApp |
Where-Object { $_.Owner.DisplayName -eq 'Name Example' } |
Select-Object DisplayName, AppName, EnvironmentName

Get-AdminDeletedPowerAppsList -EnvironmentName 'a0000000-0000-0a00-00a0-000000000000'
Get-AdminRecoverDeletedPowerApp -AppName 'AppName' -EnvironmentName 'EnvironmentName'