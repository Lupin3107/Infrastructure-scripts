# Caminho do arquivo CSV
$csvPath = "C:\Scripts\Licenses\List_Licenses.csv"

# Criação da pasta de logs
$logFolder = "C:\Logs\Licenses"
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# Importa os usuários
$users = Import-Csv -Path $csvPath

# Lista para armazenar os resultados
$resultados = @()

# Mapeamento SKUID para nomes amigáveis
$skuMap = @{
    "00a000a0-0000-a000-0000-00a0a000a000" = "Planner and Project Plan 1"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Planner and Project Plan 3"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Visio Plan 2"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Power BI PRO"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Microsoft 365 E3"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Office 365 E1 Plus"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Microsoft 365 F3"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Microsoft 365 F1"
}

foreach ($user in $users) {
    $upn = $user.UPN_encontrados
    Write-Host "`n🔍 Verificando UPN: $upn"

    try {
        $usuario = Get-MgUser -UserId $upn -ErrorAction Stop
        if ($null -eq $usuario) {
            Write-Host "! Usuário não encontrado: $upn" -ForegroundColor Yellow
            continue
        }

        # Licenças provisionadas
        $licencasDetalhadas = Get-MgUserLicenseDetail -UserId $upn -ErrorAction SilentlyContinue
        $licencas = @()


        foreach ($licenca in $licencasDetalhadas) {
            $skuId = $licenca.SkuId.ToString()
            # Only include SKUs that are present in the mapping. Skip all others.
            if ($skuMap.ContainsKey($skuId)) {
                $licencas += $skuMap[$skuId]
            } else {
                # SKUs not in $skuMap are intentionally skipped
                continue
            }
        }

        # Remove duplicates and build output text. If none matched, leave blank.
        if ($licencas.Count -gt 0) {
            $licencas = $licencas | Select-Object -Unique
            $licencasTexto = $licencas -join ', '
            Write-Host "Licenças encontradas: $licencasTexto"
        } else {
            $licencasTexto = ''
            Write-Host "Nenhuma licença registrada na variável encontrada para $upn" -ForegroundColor Yellow
        }

        $resultados += [PSCustomObject]@{
            UserPrincipalName = $upn
            Licencas          = $licencasTexto
        }

    } catch {
        Write-Host "✕ Erro ao processar UPN: $upn" -ForegroundColor Red
        Write-Host "Detalhes do erro: $_" -ForegroundColor DarkRed
    }
}

# Exporta o resultado com apenas UPN e licenças
$logPath = "$logFolder\Result_List_Licenses_in_bulk.csv"
$resultados | Export-Csv -Path $logPath -NoTypeInformation
Write-Host "✓ Processo concluído. Resultados salvos em: $logPath" -ForegroundColor Green