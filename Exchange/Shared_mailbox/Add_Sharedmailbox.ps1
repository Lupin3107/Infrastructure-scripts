# Solicita o nome da Shared Mailbox uma única vez
$sharedMailbox = Read-Host "Digite o nome da Caixa Compartilhada "

# Criação da pasta de logs
$logFolder = "C:\Example\Logs\path\Sharedmailbox"
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# Caminho do CSV contendo as permissões
$csvPath = "C:\Example\path\Add_Sharedmailbox.csv"

# Importa as contas do arquivo CSV para a variável
$entries = Import-Csv -Path $csvPath
$logs = @()

foreach ($entry in $entries) {
    $sendAsUser = $entry.SendAs_User
    $fullAccessUser = $entry.FullAccess_User

    try {
        # Adiciona permissão SendAs
        Add-RecipientPermission -Identity $sharedMailbox -AccessRights SendAs -Trustee $sendAsUser -Confirm:$False | Out-null

        # Adiciona permissão FullAccess
        Add-MailboxPermission -Identity $sharedMailbox -User $fullAccessUser -AccessRights FullAccess -InheritanceType All -Confirm:$False | Out-null

        $logs += [PSCustomObject]@{
            SharedMailbox  = $sharedMailbox
            SendAsUser     = $sendAsUser
            FullAccessUser = $fullAccessUser
            Status         = "Sucesso"
            Mensagem       = "Permissões concedidas"
        }
    }
    catch {
        $logs += [PSCustomObject]@{
            SharedMailbox  = $sharedMailbox
            SendAsUser     = $sendAsUser
            FullAccessUser = $fullAccessUser
            Status         = "Erro"
            Mensagem       = $_.Exception.Message
        }
    }
}

# Salva o log
$logPath = "C:\Example\path\Result_permissons_SharedMailbox.csv"
$logs | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

Write-Host "Processo concluído. Log salvo em: $logPath" -ForegroundColor Green
