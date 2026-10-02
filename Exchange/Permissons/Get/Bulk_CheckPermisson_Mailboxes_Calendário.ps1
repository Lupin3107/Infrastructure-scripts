$UPN = @(
"Example1@domain.com.br"
"Example2@domain.com.br"
)

$UPN | ForEach-Object {
    Get-MailboxFolderPermission -Identity "$_`:\Calendar" |
        Select-Object Identity, User, AccessRights
} | Format-Table -AutoSize