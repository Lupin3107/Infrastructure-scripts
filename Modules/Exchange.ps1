# Conecta-se ao módulo do Exchange Online e verifica se a conexão foi bem-sucedida
Write-Host "↻ Conectando ao Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false

# Verifica se o módulo está carregado e se o cmdlet Get-Mailbox está disponível
if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
    $user = (Get-ConnectionInformation).UserPrincipalName
    Write-Host "✓ Conexão realizada com sucesso com a conta: $user"
} else {
    Write-Host "✕ Falha na conexão com o Exchange Online."
}