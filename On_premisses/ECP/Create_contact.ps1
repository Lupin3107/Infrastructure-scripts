# Prompt for user input
$name = Read-Host "Enter the contact's full name"
$mail = Read-Host "Enter the contact's email address"
$displayName = Read-Host "Enter the contact's display name"
$path = Read-Host "Enter the AD path (e.g., CN=Example,DC=domain,DC=com,DC=br)"

# Create the contact in Active Directory
New-ADObject -Name $name `
             -Type Contact `
             -Path $path `
             -OtherAttributes @{'mail'=$mail; 'displayName'=$displayName}