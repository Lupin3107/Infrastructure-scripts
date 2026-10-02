$remoteDomain = "@domain.mail.onmicrosoft.com"
$csvPath = "C:\Path\Example.csv"
$users = Import-Csv $csvPath

foreach ($user in $users) {
    Enable-RemoteMailbox $user.Alias `
        -RemoteRoutingAddress ($user.Alias + $remoteDomain)
}