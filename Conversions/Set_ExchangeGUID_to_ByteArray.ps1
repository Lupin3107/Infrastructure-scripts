$users = Import-Csv "C:\Example\path\UPN_ExchGuid_extraction.csv"

foreach ($user in $users) {
    $guidByteArray = ([System.Guid]$user.ExchangeGuid).ToByteArray()
    Set-ADUser -Identity $user.UPN -Replace @{msExchMailboxGUID=$guidByteArray}
}