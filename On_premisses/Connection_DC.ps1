$cred = Get-Credential

$psSession = New-PSSession `
    -ConfigurationName Microsoft.Exchange `
    -ConnectionUri "http://DCExample.domain.com.br/PowerShell/" `
    -Authentication Kerberos `
    -Credential $cred `
    -ErrorAction Stop

Import-PSSession $psSession -DisableNameChecking -AllowClobber | Out-Null