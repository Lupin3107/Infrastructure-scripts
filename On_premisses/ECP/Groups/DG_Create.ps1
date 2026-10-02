# Prompt for values
$DisplayName = Read-Host "Enter Display Name"
$Name = Read-Host "Enter Name"
$PrimarySmtpAddress = Read-Host "Enter Primary SMTP Address (e.g., Example@domain.com.br)"

# Static OU (unchanged)
$OrganizationalUnit = "OU=Distribution Lists,OU=Example,OU=Groups,OU=Objects,OU=Objects,DC=domain,DC=com,DC=br"

# Create the distribution group
New-DistributionGroup `
    -DisplayName $DisplayName `
    -Name $Name `
    -PrimarySmtpAddress $PrimarySmtpAddress `
    -OrganizationalUnit $OrganizationalUnit