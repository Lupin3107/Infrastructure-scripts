# Get all user mailboxes with Archive disabled (server-side filter)
$users = Get-EXOMailbox -ResultSize Unlimited -Filter 'ArchiveStatus -eq "None" -and RecipientTypeDetails -eq "UserMailbox"' -Properties AccountDisabled

# Filter only active users
$activeUsers = $users | Where-Object {
    $_.AccountDisabled -ne $true
}

# Select relevant properties
$exportData = $activeUsers | Select-Object DisplayName, UserPrincipalName

# Export to CSV
$path = "C:\Logs\Planner\Archive\ArchiveDisabled_ActiveUsers.csv"
$exportData | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

Write-Host "Export completed. File saved to $path"