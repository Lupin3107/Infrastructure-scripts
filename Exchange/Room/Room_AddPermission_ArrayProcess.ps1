$rooms = @(
"Example1@domain.com.br"
"Example2@domain.com.br"
)

foreach ($room in $rooms) {
    Add-MailboxFolderPermission `
    -Identity "${room}:\Calendar" `
    -User "Example@domain.com.br" `
    -AccessRights Editor 
    | Select-Object User, AccessRights
}