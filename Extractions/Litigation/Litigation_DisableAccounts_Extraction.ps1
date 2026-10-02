# Get all mailboxes with Litigation Hold disabled
 $users = Get-Mailbox -ResultSize Unlimited | Where-Object {
     $_.LitigationHoldEnabled -eq $false -and $_.RecipientTypeDetails -eq 'UserMailbox'
 }

 # Filter only active users
 $activeUsers = $users | Where-Object {
     $_.AccountDisabled -ne $true
 }

 # Select relevant properties
 $exportData = $activeUsers | Select-Object DisplayName, UserPrincipalName

 # Export to CSV
 $exportData | Export-Csv -Path "C:\Logs\Extractions\Exchange\LitigationHoldLitigationHoldDisabled_ActiveUsers.csv" -NoTypeInformation

Write-Host "Export completed. File saved to C:\Logs\Extractions\Exchange\LitigationHoldLitigationHoldDisabled_ActiveUsers.csv"