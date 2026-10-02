# Define the path to the CSV file
$csvPath = "C:\Example\path\UPN_ExchGUID_extraction.csv"

# Import the CSV file
$users = Import-Csv -Path $csvPath

# Create a list to store results
$results = @()

# Loop through each user and get ExchangeGUID
foreach ($user in $users) {
    $mailbox = Get-Mailbox $user.UPN
    $results += [PSCustomObject]@{
        UPN       = $user.UPN
        ExchGUID  = $mailbox.ExchangeGuid
    }
}

# Export the results to the same CSV file (overwrites original)
$results | Export-Csv -Path $csvPath -NoTypeInformation