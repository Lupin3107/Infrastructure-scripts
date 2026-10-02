# Restaurar o itemclass
Get-RecoverableItems `
-Identity "Example@domain.com.br" | Where-Object { $_.ItemClass -eq "IPM.Appointment" }