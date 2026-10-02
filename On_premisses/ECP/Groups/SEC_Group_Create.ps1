$resultados = @()

# Caminho da OU
$ouPath = "OU=Groups,OU=Objects,DC=domain,DC=com,DC=br"

# Importa CSV
$groups = Import-Csv "C:\Path\Example.csv"

foreach ($group in $groups) {

    $nomeGrupo = $group.GroupName.Trim()
    $descricaoGrupo = $group.Description

    try {
        New-ADGroup -Name $nomeGrupo `
                    -SamAccountName $nomeGrupo `
                    -GroupCategory Security `
                    -GroupScope Global `
                    -Path $ouPath `
                    -Description $descricaoGrupo `
                    -ErrorAction Stop

        Write-Host "Grupo '$nomeGrupo' criado com sucesso."

        $resultados += [PSCustomObject]@{
            GroupName = $nomeGrupo
            Status    = "Criado com sucesso"
        }

    } catch {
        Write-Warning "Erro ao criar grupo '$nomeGrupo'"

        $resultados += [PSCustomObject]@{
            GroupName = $nomeGrupo
            Status    = "Erro: $($_.Exception.Message)"
        }
    }
}

# Exporta resultado
$resultados | Export-Csv "C:\Logs\Path\Example.csv" -NoTypeInformation -Encoding UTF8