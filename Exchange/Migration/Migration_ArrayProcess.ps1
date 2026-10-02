$batchName = Read-Host "Enter the migration batch name"
$csvPath   = "C:\Example\path\Migration.csv"

# 1) Create
New-MigrationBatch -Name $batchName `
  -SourceEndpoint "hybridExample@domain.com.br" `
  -CSVData ([System.IO.File]::ReadAllBytes($csvPath)) `
  -TargetDeliveryDomain "example.mail.onmicrosoft.com" `
  -AutoComplete

# 2) Set notification recipients
Set-MigrationBatch -Identity $batchName `
  -NotificationEmails "example@domain.com.br"

# 3) Start the batch explicitly
Start-MigrationBatch -Identity $batchName

Start-sleep -Seconds 15

# 4) Get the batch state
Get-MigrationBatch -Identity $batchName |  Format-List identity, state