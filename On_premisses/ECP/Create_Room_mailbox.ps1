# Prompt for user input
$name = Read-Host "Enter the room mailbox name"
$displayName = Read-Host "Enter the display name"
$alias = Read-Host "Enter the alias"

# Create the room mailbox
New-Mailbox -Room `
            -Name $name `
            -DisplayName $displayName `
            -Alias $alias `
            -Database "Example-DB"