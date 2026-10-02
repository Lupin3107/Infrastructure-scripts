<# 
    Exporta as salas (Room Mailboxes) nas quais um usuário é delegado/aprovador.
    • Funciona no Exchange Online (EXO V3).
    • Pode filtrar apenas salas que encaminham solicitações aos delegados (aprovação manual).
#>

# 2) Entradas
$approverUpn = Read-Host "Digite o UPN/e-mail do aprovador (ex.: Example@domain.com)"
$outCsv      = Read-Host "Digite o caminho do CSV de saída (ex.: C:\Path\Example.csv)"
$onlyManual  = Read-Host "Filtrar apenas salas que encaminham solicitações aos delegados? (S/N)"

# 3) Resolver o destinatário do aprovador
$approver = Get-Recipient -Identity $approverUpn -ErrorAction SilentlyContinue
if (-not $approver) {
    Write-Error "Não encontrei '$approverUpn'. Verifique o UPN/e-mail e tente novamente."
    return
}

# Preparar chaves para comparação robusta (alguns tenants retornam DN, DisplayName ou SMTP em ResourceDelegates)
$approverKeys = @(
    $approver.PrimarySmtpAddress
    $approver.Alias
    $approver.DisplayName
    $approver.LegacyExchangeDN
    $approver.Id
) | Where-Object { $_ } | ForEach-Object { $_.ToString().ToLowerInvariant() }

# 4) Buscar todas as salas
$rooms = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
         Select-Object DisplayName, PrimarySmtpAddress, Identity

# 5) Inspecionar CalendarProcessing de cada sala e filtrar por delegate
$results = foreach ($room in $rooms) {
    try {
        $cal = Get-CalendarProcessing -Identity $room.Identity -ErrorAction Stop

        # Normalizar a saída dos delegados em string para comparação
        $delegates = @($cal.ResourceDelegates) | ForEach-Object { $_.ToString().ToLowerInvariant() }

        # Checar se o aprovador está listado entre os delegados
        $isDelegate = $false
        foreach ($k in $approverKeys) {
            if ($delegates -contains $k -or ($delegates | Where-Object { $_ -like "*$k*" })) {
                $isDelegate = $true
                break
            }
        }

        if (-not $isDelegate) { continue }

        # Se o usuário pediu "apenas salas com aprovação manual", aplicar lógica adicional:
        $passesManualFilter = $true
        if ($onlyManual -match '^[sS]') {
            # Consideramos aprovação manual quando a sala encaminha solicitações aos delegados
            # e não aceita tudo silenciosamente.
            $passesManualFilter = $cal.ForwardRequestsToDelegates -eq $true
        }

        if ($passesManualFilter) {
            [PSCustomObject]@{
                Sala                  = $room.DisplayName
                EmailSala             = $room.PrimarySmtpAddress
                Aprovador             = $approver.PrimarySmtpAddress
                EncaminhaParaDelegado = $cal.ForwardRequestsToDelegates
                AutomateProcessing    = $cal.AutomateProcessing
                Delegados             = ($cal.ResourceDelegates -join '; ')
            }
        }
    }
    catch {
        Write-Warning "Falha ao ler '$($room.DisplayName)': $($_.Exception.Message)"
    }
}

# 6) Exportar
if ($results -and $results.Count -gt 0) {
    $results | Sort-Object Sala | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Concluído! $($results.Count) sala(s) exportada(s) em: $outCsv" -ForegroundColor Green
} else {
    Write-Host "Nenhuma sala encontrada para o aprovador informado com os filtros aplicados." -ForegroundColor Yellow
}

# 7) Encerrar sessão (opcional)
# Disconnect-ExchangeOnline -Confirm:$false