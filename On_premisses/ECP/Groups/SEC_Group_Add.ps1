$resultados = @()

# Solicita o nome do grupo ao usuário
$grupo = Read-Host "Digite o nome do grupo"

# Adiciona membros ao grupo
Import-Csv "C:\Path\Example.csv" | ForEach-Object {
    $UPN = $_.UPN.Trim()
    if (![string]::IsNullOrWhiteSpace($UPN)) {
        try {
            Add-ADGroupMember -Identity $grupo -Members $UPN -ErrorAction Stop
            Write-Host "Adicionado: $UPN"
            $resultados += [PSCustomObject]@{
                UPN = $UPN
                Status = "Adicionado com sucesso"
            }
        } catch {
            Write-Warning "Erro ao adicionar $UPN"
            $resultados += [PSCustomObject]@{
                UPN = $UPN
                Status = "Erro: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Warning "Linha ignorada: UPN vazio ou inválido."
        $resultados += [PSCustomObject]@{
            UPN = $UPN
            Status = "Ignorado: UPN vazio ou inválido"
        }
    }
}

# Exporta os resultados para CSV
$resultados | Export-Csv "C:\Logs\Path\Example.csv" -NoTypeInformation -Encoding UTF8