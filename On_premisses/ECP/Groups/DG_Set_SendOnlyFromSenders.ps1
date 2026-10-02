$dg = "Example@Domain.com.br"
(Get-DistributionGroup -Identity $dg).AcceptMessagesOnlyFromSendersOrMembers| Select Name

$alias = (Import-Csv C:\Logs\leonardosouza.idr\Delivery_management\DM_ECP_Alias.csv).Alias
Set-DistributionGroup -Identity $dg -AcceptMessagesOnlyFromSendersOrMembers $Alias