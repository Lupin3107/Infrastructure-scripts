# extrair qual é o nome do calendário
Get-EXOMailboxFolderStatistics `
    -Identity "Example@domain.com.br" |
Where-Object {
    $_.FolderType -match "Calendar"
} | Select-Object Name

# extrair os logs do calendário para análise dos dados
Get-CalendarDiagnosticObjects `
-identity "Example@domain.com.br" `
-ShouldFetchRecurrenceExceptions $false | Where-Object { $_.IsException -eq $false -and $_.SubjectProperty -eq "Subject Example" }
| Select-Object LogTimestamp,CalendarLogTriggerAction,LogClientInfoString,ItemClass,ParentDisplayName,OriginalParentDisplayName,IsException,IsRecurring,IsSeriesCancelled,MeetingRequestType,Sender,ResponsibleUserName,RecurrencePattern,SubjectProperty
