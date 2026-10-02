# SKU único a ser removido
$TargetSkuId = "00a000a0-0000-a000-0000-00a0a000a000"  # Copilot Pro

# Pasta e arquivo de log
$LogFolder = "C:\Logs\Licenses\Copilot"
if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}
$LogFile = Join-Path $LogFolder "Copilot_remocao_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$Encoding = "UTF8Bom"

function Write-Log {
    param(
        [Parameter(Mandatory)][pscustomobject]$Row
    )
    if (-not (Test-Path -LiteralPath $LogFile)) {
        $Row | Export-Csv -Path $LogFile -Encoding $Encoding -NoTypeInformation
    } else {
        $Row | Export-Csv -Path $LogFile -Encoding $Encoding -Append
    }
}

# Fonte do CSV
$CSV_Source_Path = "C:\Scripts\Licenses\Licenses_to_add.csv"

if (-not (Test-Path $CSV_Source_Path)) {
    Write-Error "Arquivo CSV não encontrado em: $CSV_Source_Path"
    exit 1
}

$Users = Import-Csv -Path $CSV_Source_Path

foreach ($user in $Users) {
    $upn = $user.Copilot
    if ([string]::IsNullOrWhiteSpace($upn)) {
        continue
    }

    # Verifica se o usuário existe
    try {
        $mgUser = Get-MgUser -UserId $upn -ErrorAction Stop
        Write-Host "Usuário(a) verificado(a) ✓ $upn"
    } catch {
        Write-Warning "Usuário(a) não encontrado(a) ✕"
        $LogRow = [pscustomobject]@{
            UPN          = $upn
            Acao         = "Verificacao"
            Licenca_Alvo = $TargetSkuId
            Resultado    = "Erro"
            Detalhe      = "Usuário(a) não encontrado(a)"
            Executado    = "Falso"
            Autor        = $env:USERNAME
            DataHora     = (Get-Date)
        }
        Write-Log -Row $LogRow
        continue
    }

    try {
        # Licenças atuais do usuário
        $licensesAssigned = (Get-MgUserLicenseDetail -UserId $upn | Select-Object -ExpandProperty SkuId)

        # Se não possui a licença alvo, não faz nada
        if ($licensesAssigned -notcontains $TargetSkuId) {
            Write-Host "`tNenhuma ação necessária. Não possui a licença alvo 🧾"
            $LogRow = [pscustomobject]@{
                UPN          = $upn
                Acao         = "Remocao"
                Licenca_Alvo = $TargetSkuId
                Resultado    = "SemAlteracao"
                Detalhe      = "Usuário(a) não possui a licença"
                Executado    = "Verdadeiro"
                Autor        = $env:USERNAME
                DataHora     = (Get-Date)
            }
            Write-Log -Row $LogRow
            continue
        }

        # Remove APENAS a licença alvo, sem adicionar nenhuma outra
        Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses @($TargetSkuId) | Out-Null

        Write-Host "`tLicença removida 🗑️"
        $LogRow = [pscustomobject]@{
            UPN          = $upn
            Acao         = "Remocao"
            Licenca_Alvo = $TargetSkuId
            Resultado    = "Sucesso"
            Detalhe      = "Licença removida"
            Executado    = "Verdadeiro"
            Autor        = $env:USERNAME
            DataHora     = (Get-Date)
        }
        Write-Log -Row $LogRow

    } catch {
        Write-Warning "Falha ao remover licença ✕"
        Write-Host $_.Exception.Message

        $LogRow = [pscustomobject]@{
            UPN          = $upn
            Acao         = "Remocao"
            Licenca_Alvo = $TargetSkuId
            Resultado    = "Erro"
            Detalhe      = $_.Exception.Message
            Executado    = "Falso"
            Autor        = $env:USERNAME
            DataHora     = (Get-Date)
        }
        Write-Log -Row $LogRow
    }
}