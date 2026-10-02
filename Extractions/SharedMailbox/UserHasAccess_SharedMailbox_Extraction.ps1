
 # Specify the user to search for
$UserAccount = Read-host "Enter the user account"
 
# Get all shared mailboxes
$SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo #| Where-Object {$_.RecipientTypeDetails -ne "DiscoveryMailbox"} 
 
# Find shared mailboxes where the user has access
$AccessList = foreach ($Mailbox in $SharedMailboxes) {
   Write-Host -f Cyan "Checking mailbox: $($Mailbox.PrimarySmtpAddress)..."
   # Full Access
   $FullAccess = Get-EXOMailboxPermission -Identity $Mailbox.Identity  | 
      Where-Object { $_.User -contains $UserAccount -and $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false }
   If($FullAccess) {
      Write-Host -f Green "`tFound FullAccess permission for $UserAccount on mailbox $($Mailbox.PrimarySmtpAddress)."
      $FullPermissions = "Yes"
   }
   else {
      $FullPermissions = "No"
   }
 
   # Send As
   $SendAs = Get-EXORecipientPermission -Identity $Mailbox.Identity | Where-Object { $_.Trustee -contains $UserAccount -and $_.AccessRights -contains "SendAs" }
   If($SendAs) {
      Write-Host -f Green "`tFound SendAs permission for $UserAccount on mailbox $($Mailbox.PrimarySmtpAddress)."
      $SendAsPermissions = "Yes"
   }
   Else {
      $SendAsPermissions = "No"
   }
    
   # Send On Behalf
   $Recipient = Get-ExoRecipient -Identity $UserAccount
   $SendOnBehalf = $Mailbox.GrantSendOnBehalfTo | Where-Object {$_ -contains $Recipient.Identity}
   If($SendOnBehalf) {
     Write-Host -f Green "`tFound SendOnBehalf permission for $UserAccount on mailbox $($Mailbox.PrimarySmtpAddress)."
     $SendOnBehalfPermissions = "Yes"
   }
   Else {
     $SendOnBehalfPermissions = "No"
   }
    
   # If any permissions found, add to access list
   if ($FullPermissions -eq "Yes" -or $SendAsPermissions -eq "Yes" -or $SendOnBehalfPermissions -eq "Yes") {
      [PSCustomObject]@{
         SharedMailbox = $Mailbox.DisplayName
         PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
         FullPermissions = $FullPermissions
         SendAsPermissions = $SendAsPermissions
         SendOnBehalfPermissions = $SendOnBehalfPermissions
      }
   }
}
 
# Display results
$AccessList | Format-Table -AutoSize
 
# Export to CSV (optional)
$AccessList | Export-Csv -Path "C:\Logs\Exchange\Shared_Mailbox\User_AccessReport_$UserAccount.csv" -NoTypeInformation