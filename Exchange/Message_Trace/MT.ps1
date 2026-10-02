$StartDate = Get-Date "2026-05-01 00:00:00"
$EndDate   = Get-Date "2026-07-06 00:00:00"

Get-MessageTraceV2 `
-StartDate $StartDate `
-EndDate $EndDate `
-ResultSize 5
-SenderAddress "example1@domain.com" `
-Subject `
-SubjectFilterType Contains `
-RecipientAddress "example2@domain.com"
-MessageId