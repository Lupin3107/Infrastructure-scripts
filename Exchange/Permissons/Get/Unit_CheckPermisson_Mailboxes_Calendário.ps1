$targetUser = 'NAMES'
$UPN = @(
    "Example@domain.com.br"
)
$UPN | ForEach-Object {
    $current = $_
    Get-MailboxFolderPermission -Identity "$current`:\Calendário" |
        Where-Object {
        $_.User.ToString() -eq $targetUser
        } |
        Select-Object @{n='UPN';e={$current}}, User, AccessRights
} | Format-Table -AutoSize