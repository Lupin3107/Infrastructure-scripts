$targetUpn = "example2@domain.com"

# Find inactive mailbox
$inactive = Get-Mailbox -InactiveMailbox -Identity "example1@domain.com"

# Restoring inactive mailbox
Restore-InactiveMailbox `
-InactiveMailbox "example@domain.com" `
-TargetMailbox $targetUpn