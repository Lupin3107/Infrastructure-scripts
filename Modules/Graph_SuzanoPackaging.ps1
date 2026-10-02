# Conecta-se ao modulo do MS Graph necessário para executar o código e verifica se a conexão foi bem-sucessedida
Write-Host "↻ Conectando ao MS Graph..."
Connect-MgGraph -NoWelcome  -tenantID "00a000a0-0000-a000-0000-00a0a000a000" -ContextScope Process -Scopes "User.ReadWrite.All"

$context = Get-MgContext
if ($context.Account -and $context.Scopes) {
    Write-Host "✓ Conexão realizada com sucesso com a conta: $($context.Account)"
} else {
    Write-Host "✕ Falha na conexão com o Microsoft Graph."
}