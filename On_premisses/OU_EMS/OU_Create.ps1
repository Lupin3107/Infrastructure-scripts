    # Prompt for OU Name
$OUName = Read-Host "Enter OU Name"

# Static parent OU (unchanged)
$ParentOU = "OU=Objetos,DC=domain,DC=com,DC=br"

# Create the OU
New-ADOrganizationalUnit `
    -Name $OUName `
    -Path $ParentOU