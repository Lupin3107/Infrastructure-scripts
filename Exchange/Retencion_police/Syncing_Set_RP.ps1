$csvPath = "C:\Example\path\iniciar_RP.csv"
$logPath = "C:\Example\path\Iniciar_RP_Results.csv"

# Validar CSV
if (-not (Test-Path $csvPath)) {
    Write-Host "CSV não encontrado: $csvPath" -ForegroundColor Red
    exit 1
}

$users = Import-Csv $csvPath

if (-not ($users[0].PSObject.Properties.Name -contains 'UPN')) {
    Write-Host "CSV precisa conter a coluna 'UPN'" -ForegroundColor Red
    exit 1
}

# Garantir pasta de log
$dir = Split-Path $logPath
if (-not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

Write-Host "`r`n Executando Managed Folder Assistant para $($users.Count) usuários..." -ForegroundColor Cyan
Write-Host ""

$index = 0

$results = $users | ForEach-Object {
    $index++
    $upn = $_.UPN.Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Host ("[{0}/{1}] UPN vazio -> ignorado" -f $index, $users.Count) -ForegroundColor DarkYellow

        return [pscustomobject]@{
            UPN    = ""
            Status = "Failed"
            Error  = "Blank UPN"
        }
    }

    Write-Host ("[{0}/{1}] Processando: {2}" -f $index, $users.Count, $upn) -ForegroundColor Yellow

    try {
        Start-ManagedFolderAssistant -Identity $upn -ErrorAction Stop

        Write-Host "   ✔ Sucesso" -ForegroundColor Green

        [pscustomobject]@{
            UPN    = $upn
            Status = "Success"
            Error  = ""
        }
    }
    catch {
        Write-Host ("✕ Erro: {0}" -f $_.Exception.Message) -ForegroundColor Red

        [pscustomobject]@{
            UPN    = $upn
            Status = "Failed"
            Error  = $_.Exception.Message
        }
    }
}

# Exportar resultado
$results | Export-Csv $logPath -NoTypeInformation -Encoding UTF8

Write-Host "`r`n Finalizado. Resultado em: $logPath" -ForegroundColor Cyan

# Resumo rápido
$success = ($results | Where-Object Status -eq "Success").Count
$failed  = ($results | Where-Object Status -eq "Failed").Count

Write-Host "`r`n Resumo -> Sucesso: $success | Falha: $failed" -ForegroundColor White