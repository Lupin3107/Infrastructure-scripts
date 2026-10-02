$resultados = @()

# Solicita o nome do grupo ao usuário
$grupo = Read-Host "Digite o e-mail do grupo (ex: Group@domain.com.br)"

Import-Csv "C:\Path\Example.csv" | ForEach-Object {
    $email = $_.email.Trim()
    if (![string]::IsNullOrWhiteSpace($email)) {
        try {
            Add-DistributionGroupMember -Identity $grupo -Member $email -ErrorAction Stop
            Write-Host "Adicionado: $email"
            $resultados += [PSCustomObject]@{
                Email = $email
                Status = "Adicionado com sucesso"
            }
        } catch {
            Write-Warning "Erro ao adicionar $email"
            $resultados += [PSCustomObject]@{
                Email = $email
                Status = "Erro: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Warning "Linha ignorada: campo de e-mail vazio ou inválido."
        $resultados += [PSCustomObject]@{
            Email = $email
            Status = "Ignorado: e-mail vazio ou inválido"
        }
    }
}

# Exporta os resultados para CSV
$resultados | Export-Csv "C:\Logs\Path\Example.csv" -NoTypeInformation -Encoding UTF8