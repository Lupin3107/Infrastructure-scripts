# Path to CSV with UPN and Phone columns
$csvPath = "C:\Scripts\Temp\Temp.csv"
$users = Import-Csv -Path $csvPath

# Ensure you have the Microsoft Graph PowerShell modules installed and admin consent for the scopes below
# Install-Module Microsoft.Graph -Scope CurrentUser
# You need at least: UserAuthenticationMethod.ReadWrite.All and User.ReadWrite.All
# Connect interactively (this will prompt for login)
Connect-MgGraph -Scopes "UserAuthenticationMethod.ReadWrite.All","User.ReadWrite.All"

$success = 0
$skipped = 0
$failed = 0

foreach ($user in $users) {
    $upn = $user.UPN.Trim()
    $phone = if ($user.Phone) { $user.Phone.Trim() } else { $null }

    if (-not $upn) {
        Write-Warning "Skipping row with empty UPN"
        $skipped++
        continue
    }

    if (-not $phone) {
        Write-Warning "Skipping $upn — no phone provided"
        $skipped++
        continue
    }

    try {
        # Verify user exists
        $mgUser = Get-MgUser -UserId $upn -ErrorAction Stop
    } catch {
        Write-Warning "User $upn not found"
        $failed++
        continue
    }

    try {
        # Try to create an authentication phone method (MFA phone)
        New-MgUserAuthenticationPhoneMethod -UserId $mgUser -Phone $phone -PhoneType "mobile" -ErrorAction Stop
        Write-Host "Added auth phone method for $upn -> $phone" -ForegroundColor Green
        $success++
    } catch {
        # If creation fails (often because a method already exists), attempt to update an existing method
        Write-Warning "Create failed for $upn $_. Trying to update existing method..."
        try {
            $methods = Get-MgUserAuthenticationPhoneMethod -UserId $upn -ErrorAction Stop
            # Prefer an existing mobile method, otherwise take the first available
            $method = $methods | Where-Object { $_.PhoneType -eq 'mobile' } | Select-Object -First 1
            if (-not $method) { $method = $methods | Select-Object -First 1 }

            if ($method) {
                Update-MgUserAuthenticationPhoneMethod -UserId $upn -AuthenticationPhoneMethodId $method.Id -Phone $phone -ErrorAction Stop
                Write-Host "Updated auth phone method for $upn -> $phone" -ForegroundColor Yellow
                $success++
            } else {
                Write-Warning "No phone methods found to update for $upn"
                $failed++
            }
        } catch {
            Write-Warning "Failed to update phone for $upn $_"
            $failed++
        }
    }

    # ALTERNATIVE: If you prefer to set the Azure AD 'mobilePhone' attribute (visible on user object) instead of the authentication method,
    # uncomment the next line and comment out the New/Update blocks above.
    # Update-MgUser -UserId $upn -MobilePhone $phone -ErrorAction SilentlyContinue
}

Write-Host "\nSummary: Success=$success, Skipped=$skipped, Failed=$failed" -ForegroundColor Cyan