# Mapeando os SkuIDs e nomeando-os
$SkuIDsToAdd = @(
    "00a000a0-0000-a000-0000-00a0a000a000"  # Office 365 E1_Plus
)
$SkuIDsToRemove = @(
    "00a000a0-0000-a000-0000-00a0a000a000",  # Microsoft 365 E3
    "00a000a0-0000-a000-0000-00a0a000a000",  # Microsoft 365 F3
    "00a000a0-0000-a000-0000-00a0a000a000",  # Microsoft 365 F1
    "00a000a0-0000-a000-0000-00a0a000a000",  # Exchange Kiosk
    "00a000a0-0000-a000-0000-00a0a000a000"   # Exchange Archiving
)

# Criando as variáveis para a função de Log
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Row
    )
    if (-not (Test-Path -LiteralPath $LogFile)) {
        # Primeira escrita: criar com cabeçalho (sem -Append)
        $Row | Export-Csv -Path $LogFile -Encoding $Encoding -NoTypeInformation
    } else {
        # Escritas subsequentes: anexar ao arquivo existente
        $Row | Export-Csv -Path $LogFile -Encoding $Encoding -Append
    }
}

$LogFolder = "C:\Logs\Licenses\E1"
if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}
$LogFile = Join-Path $LogFolder "E1_atribuicao_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$Encoding = "UTF8Bom"

# Definindo as variáveis para a origem do CSV e importando
$CSV_Source_Path = "C:\Scripts\Licenses\Licenses_to_add.csv"
$Users = Import-Csv -Path $CSV_Source_Path

foreach ($user in $Users) {
    $upn = $user.E1
    try {
        $existingUser = Get-MgUser -UserId $upn
    } catch {
        $existingUser = $null
        }

    if ($existingUser) {
        Write-Host "Usuário(a) verificado(a) ✓ $upn |" -NoNewline
    } else {
        Write-Host "Usuário(a) não encontrado(a) ✕ |" -NoNewline
            
        $LogUserNotFound = [pscustomobject]@{
            UPN        = $upn
            Erro       = "Usuário(a) não encontrado(a)"
            Executado  = "Falso"
            Autor      = $env:USERNAME
        }
        
        Write-Log -Row $LogUserNotFound
        continue
    }
    try { # Conferindo as licenças que precisam ser removidas e adicionadas para aplicar as mudanças ao usuário
        
        # Interseção entre as licenças a remover/adicionar com as licenças atribuídas ao usuário
        $licensesAssigned = (Get-MgUserLicenseDetail -UserId $upn).SkuId
        $SkuIDsToRemove_intersection = $licensesAssigned | Where-Object { $_ -in $SkuIDsToRemove }
        $SkuIDsToAdd_intersection = $SkuIDsToAdd | Where-Object { $licensesAssigned -contains $_ }
        
        # Conferindo se as licenças a adicionar já estão atribuídas; se sim, não tentaremos adicioná-las novamente
        if ($SkuIDsToAdd_intersection.Count -gt 0) {
            $SkuIDs_Checked = $SkuIDsToAdd | Where-Object { $licensesAssigned -notcontains $_ }
        } else {
            $SkuIDs_Checked = $SkuIDsToAdd
        }
        
        # Converter as licenças verificadas para o formato compatível com o parâmetro -AddLicenses
        $AddLicenses_Checked = $SkuIDs_Checked | ForEach-Object { @{ SkuId = $_ } }
        # Preparar os SkuIDs em formato legível para registro no CSV
        $SkuID_csv_string_LicensesAdd = $AddLicenses_Checked | ForEach-Object { $_.SkuId }

        # Tentar adicionar e remover as licenças do usuário, se necessário
        if ((@($AddLicenses_Checked).Count + @($SkuIDsToRemove_intersection).Count) -gt 0) {
            Set-MgUserLicense -UserId $upn `
                -AddLicenses $AddLicenses_Checked `
                -RemoveLicenses @($SkuIDsToRemove_intersection) | Out-Null
            
            $LogLicensesUpdated = [pscustomobject]@{
                UPN            = $upn
                Licenças_adicionadas = if ([string]::IsNullOrWhiteSpace(($SkuID_csv_string_LicensesAdd -join ';'))) {
                    "Não foi necessário adicionar licenças"
                } else {
                    $SkuID_csv_string_LicensesAdd -join '/'
                }
                Licenças_removidas = if ([string]::IsNullOrWhiteSpace(($SkuIDsToRemove_intersection -join ';'))) {
                    "Não foi necessário remover licenças"
                } else {
                    $SkuIDsToRemove_intersection -join '/'
                }
                Executado     = "Verdadeiro"
                Autor         = $env:USERNAME
            }
        
            Write-Host " Licenças atualizadas 📓"
            Write-Log -Row $LogLicensesUpdated

        } else {
            $LogNoChangesNeeded = [pscustomobject]@{
                UPN            = $upn
                Licenças_adicionadas = "Não foi necessário adicionar licenças"
                Licenças_removidas = "Não foi necessário remover licenças"
                Executado     = "Verdadeiro"
                Autor         = $env:USERNAME
            }
        
            Write-Host " Licenças não atualizadas, pois o colaborador(a) já possui as mesmas 🧾"
            Write-Log -Row $LogNoChangesNeeded
        }
    } catch {
        Write-Host " Falha ao adicionar/remover licenças ✕"
        Write-Host $_.Exception.Message

        $LogException = [pscustomobject]@{
            UPN       = $upn
            Erro      = $_.Exception.Message
            Executado = "Falso"
            Autor     = $env:USERNAME
        }
        
        Write-Log -Row $LogException
    }
}