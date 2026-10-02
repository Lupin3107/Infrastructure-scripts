# Cancel all meetings organized by a deleted user on the target range
Remove-CalendarEvents `
-Identity "example@domain.com" `
-CancelOrganizedMeetings `
-QueryStartDate (Get-Date) `
-QueryWindowInDays 99999