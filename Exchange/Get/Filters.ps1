#EXO - Export | UserMailbox
Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox |
    Select-Object UserPrincipalName |
    Export-Csv -Path "C:\logs\EXO_Export_UPN.csv" -NoTypeInformation -Encoding UTF8

#ECP - Export | RemoteUserMailbox
Get-Recipient -ResultSize 5 -RecipientTypeDetails RemoteUserMailbox |
    Select-Object Name, primarysmtpaddress |
    Export-Csv -Path "C:\logs\Myfolder\ECP_Export_RemoteUserMailbox_PrimarySMTP.csv" -NoTypeInformation -Encoding UTF8

#ECP - Export | Usermailbox
Get-Recipient -ResultSize Unlimited -RecipientTypeDetails Usermailbox |
     Select-Object PrimarySmtpAddress |
     Export-Csv -Path "C:\logs\Myfolder\ECP_Export_Usermailbox_PrimarySMTP.csv" -NoTypeInformation -Encoding UTF8

<# EXO | Mailbox Properties
   - View ExchangeGUid set (Cloud)
#>
Get-Mailbox "Exemplo@suzano.com.br" | ft Identity,ExchangeGuid

<# ECP | Mailbox Properties
   - View ExchangeGUid set (On-prem)
#>
Get-RemoteMailbox "example@domain.com" 
| Select-Object Identity, ExchangeGuid, LegacyExchangeDNf

# EXO | Filter mail Mailbox
Get-EXORecipient -Filter "EmailAddresses -like '*_@domain.com*'"

# EXO | Filter size Mailbox
$UPN = @(
    "example1@domain.com",
    "example2@domain.com",
)

$UPN | ForEach-Object {
    Get-MailboxStatistics -Identity $_ |
    Select-Object DisplayName, TotalItemSize
} | Format-Table -AutoSize
