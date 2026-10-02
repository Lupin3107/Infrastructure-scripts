# Conceder permissão de Leitor
Add-MailboxFolderPermission `
-Identity "${MailboxOrigem}:\Inbox" `
-User "Example@domain.com" `
-AccessRights Reviewer
