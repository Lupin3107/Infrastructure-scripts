# Caminho da planilha (ajuste se necessário)
$csvPath = "C:\Example\path\Add_RP.csv"

# Nome da política de retenção que será aplicada
$retentionPolicy = "Example"

# Pasta de logs
$logFolder = "C:\Example\path\"
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# Importar os usuários do CSV (verifica existência do arquivo)
if (-not (Test-Path -Path $csvPath)) {
    Write-Host "Arquivo CSV não encontrado: $csvPath" -ForegroundColor Red
    return
}

$users = Import-Csv -Path $csvPath

# Coleção de resultados (UPN, Status, Message)
$results = @()

# Função para tentar extrair UPN da linha do CSV a partir de várias colunas possíveis
function Get-UPNFromRow {
    param($row)
    $possibleNames = @('UPN_encontrados','UPN','UserPrincipalName','Email','EmailAddress','mail','User')
    $props = $row.PSObject.Properties.Name
    foreach ($name in $possibleNames) {
        if ($props -contains $name) {
            $val = $row.$name
            if ($null -ne $val -and $val.ToString().Trim() -ne '') { return $val.ToString().Trim() }
        }
    }
    # Se não encontrou, tenta pegar a primeira propriedade não vazia
    foreach ($p in $props) {
        $val = $row.$p
        if ($null -ne $val -and $val.ToString().Trim() -ne '') { return $val.ToString().Trim() }
    }
    return $null
}

# Aplicar a política de retenção para cada usuário, registrando sucesso/falha
foreach ($user in $users) {
    $upn = Get-UPNFromRow -row $user
    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Host "! UPN não encontrado na linha do CSV. Pulando." -ForegroundColor Yellow
        $results += [pscustomobject]@{
            UPN     = ''
            Status  = 'Failed'
            Message = 'UPN not found in CSV row'
        }
        continue
    }

    try {
        Write-Host "Aplicando política '$retentionPolicy' para: $upn" -ForegroundColor yellow
        # Use -ErrorAction Stop para capturar falhas
        Set-Mailbox -Identity $upn -RetentionPolicy $retentionPolicy -ErrorAction Stop
        Write-Host "✔ Política aplicada com sucesso para: $upn" -ForegroundColor Green
        $results += [pscustomobject]@{
            UPN     = $upn
            Status  = 'Success'
            Message = ''
        }
    } catch {
        $errMsg = $_.Exception.Message
        Write-Host "✕ Erro ao aplicar política para: $upn" -ForegroundColor Red
        Write-Host $errMsg -ForegroundColor DarkRed
        $results += [pscustomobject]@{
            UPN     = $upn
            Status  = 'Failed'
            Message = $errMsg
        }
    }
}

# Exportar resultados para CSV com timestamp na pasta de logs
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outPath = Join-Path -Path $logFolder -ChildPath "Add_Retention_policy_Results_$timestamp.csv"
$results | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8

# Resumo
$total = $results.Count
$success = ($results | Where-Object { $_.Status -eq 'Success' }).Count
$failed = ($results | Where-Object { $_.Status -eq 'Failed' }).Count

Write-Host "Resumo: Total=$total, Sucesso=$success, Falha=$failed" -ForegroundColor Cyan
Write-Host "Resultados exportados para: $outPath" -ForegroundColor Yellow

# Observação: este script requer sessão adequada para executar Set-Mailbox (Exchange Management Shell / Exchange Online PowerShell).