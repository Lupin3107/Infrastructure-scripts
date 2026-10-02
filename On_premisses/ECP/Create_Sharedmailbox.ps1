# Prompt for user input
$name = Read-Host "Enter the shared mailbox name"
$displayName = Read-Host "Enter the display name"
$alias = Read-Host "Enter the alias"

# Create the shared mailbox
New-Mailbox -Shared `
            -Name $name `
            -DisplayName $displayName `
            -Alias $alias `
            -Database "Example-DB"
