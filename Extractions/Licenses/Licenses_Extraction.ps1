# Define the target SKU IDs and their friendly names
$skuNames = @{
    "639dec6b-bb19-468b-871c-c5c441c4b0cb" = "Copilot PRO"
}

$targetSkuIds = $skuNames.Keys

# Initialize an array to store matching users
$matchedUsers = @()

# Get all users (paged)
$users = Get-MgUser -All

foreach ($user in $users) {
    try {
        # Get license details for each user
        $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id

        # Check if the user has at least one of the target licenses
        foreach ($license in $licenseDetails) {
            if ($targetSkuIds -contains $license.SkuId) {
                $matchedUsers += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    LicenseName = $skuNames[$license.SkuId]
                }
                break
            }
        }
    } catch {
        Write-Warning "Failed to get license details for user: $($user.UserPrincipalName)"
    }
}

# Export to CSV
$exportPath = "C:\logs\M365\Licensed_Users_Extraction.csv"
$matchedUsers | Export-Csv -Path $exportPath -NoTypeInformation

Write-Host "✓ Export complete. File saved at: $exportPath"