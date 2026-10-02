# Grant FullAccess permission to the target mailbox
Add-MailboxPermission -Identity "example@domain.com" `
-User "example@domain.com" `      
-AccessRights FullAccess `
-AutoMapping:$true