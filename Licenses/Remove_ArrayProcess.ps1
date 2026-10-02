# Caminho do arquivo CSV que lerá a coluna 'Remove'
$csvPath = "C:\Scripts\Licenses\List_Remove_bulk.csv"

# Criação da pasta C:\Logs\Licenses\Remove_Licenses
$logFolder = "C:\Logs\Licenses"
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory
}

# Importa as contas do arquivo CSV para a variável
$users = Import-Csv -Path $csvPath

if (-not $users -or -not ($users | Get-Member -Name Remove -MemberType NoteProperty)) {
    Write-Error "O CSV deve conter a coluna 'Remove'."
    exit 1
}

Write-Host "Processando $($users.Count) usuários para remoção total de licenças..." -ForegroundColor Cyan

# Lista de resultados para log
$logs = @()

$index = 0
foreach ($user in $users) {
    $index++
    $upn = ($user.Remove).Trim()
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Host ("[{0}/{1}] UPN em branco - ignorado" -f $index, $users.Count) -ForegroundColor DarkYellow

        $logs += [PSCustomObject]@{
            Index              = $index
            Timestamp          = $timestamp
            UserPrincipalName  = ""
            Status             = "Ignorado"
            Mensagem           = "UPN em branco"
        }
        continue
    }

    Write-Host ("[{0}/{1}] Removendo licenças de: {2}" -f $index, $users.Count, $upn) -ForegroundColor Yellow

    try {
        # Obtém licenças atribuídas ao usuário
        $licenseDetails = Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop
        $assignedSkuIds = $licenseDetails.SkuId | Where-Object { $_ } | Select-Object -Unique

        if (-not $assignedSkuIds -or $assignedSkuIds.Count -eq 0) {
            Write-Host "   Usuário sem licenças atribuídas" -ForegroundColor DarkYellow

            $logs += [PSCustomObject]@{
                Index              = $index
                Timestamp          = $timestamp
                UserPrincipalName  = $upn
                Status             = "Sem alteração"
                Mensagem           = "Usuário sem licenças atribuídas"
            }
            continue
        }

        # Remove todas as licenças
        Set-MgUserLicense -UserId $upn `
            -AddLicenses @() `
            -RemoveLicenses $assignedSkuIds | Out-Null

        Write-Host "   Licenças removidas com sucesso" -ForegroundColor Green

        $logs += [PSCustomObject]@{
            Index              = $index
            Timestamp          = $timestamp
            UserPrincipalName  = $upn
            Status             = "Sucesso"
            Mensagem           = "Licenças removidas: $($assignedSkuIds -join ', ')"
        }
    }
    catch {
        Write-Warning ("   Erro ao processar {0}: {1}" -f $upn, $_.Exception.Message)

        $logs += [PSCustomObject]@{
            Index              = $index
            Timestamp          = $timestamp
            UserPrincipalName  = $upn
            Status             = "Erro"
            Mensagem           = $_.Exception.Message
        }

        Start-Sleep -Milliseconds 200
    }
}

# Timestamp para o nome do arquivo de log
$fileTimestamp = Get-Date -Format "dd-MM-yyyy_HH-mm"

# Exporta log com timestamp no nome do arquivo
$logPath = "$logFolder\Result_Removed_All_Licenses_$fileTimestamp.csv"
$logs | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

Write-Host "Processo finalizado. Log salvo em: $logPath" -ForegroundColor Gray