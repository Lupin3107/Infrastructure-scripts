$csvPath = "C:\Example\path\MFA_Add_Phone.csv"
$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    $upn = $user.UPN
    $phone = $user.Phone

    try { # Verify user exists
    Get-MgUser -UserId $upn -ErrorAction Stop
    }catch {
        Write-Warning "Failed on verifying the user for $upn"
        }

    try { # Try to create an authentication phone method (MFA phone)
    New-MgUserAuthenticationPhoneMethod -UserId $upn -PhoneNumber $phone -PhoneType "mobile"
    }catch {
        Write-Warning "Failed on adding the phone number for the $upn"
    }
}