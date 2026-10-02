# Create the Restore
New-MailboxRestoreRequest `
-SourceMailbox "example@domain.com" `
-TargetMailbox "bkp_example@domain.com" `
-AllowLegacyDNMismatch

# See the progress
Get-MailboxRestoreRequestStatistics -Identity "bkp_example"