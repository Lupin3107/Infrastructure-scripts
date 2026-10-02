# SKU único a ser atribuído
$TargetSkuId = "00a000a0-0000-a000-0000-00a0a000a000"  # Copilot Pro

# Pasta e arquivo de log
$LogFolder = "C:\Logs\Licenses\Copilot"
if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}
$LogFile = Join-Path $LogFolder "Copilot_atribuicao_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
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
            UPN                   = $upn
            Acao                  = "Verificacao"
            Licenca_Alvo          = $TargetSkuId
            Resultado             = "Erro"
            Detalhe               = "Usuário(a) não encontrado(a)"
            Executado             = "Falso"
            Autor                 = $env:USERNAME
            DataHora              = (Get-Date)
        }
        Write-Log -Row $LogRow
        continue
    }

    try {
        # Licenças atuais do usuário
        $licensesAssigned = (Get-MgUserLicenseDetail -UserId $upn | Select-Object -ExpandProperty SkuId)

        # Se já possui a licença alvo, não faz nada
        if ($licensesAssigned -contains $TargetSkuId) {
            Write-Host "`tNenhuma ação necessária. Já possui a licença alvo 🧾"
            $LogRow = [pscustomobject]@{
                UPN                   = $upn
                Acao                  = "Atribuicao"
                Licenca_Alvo          = $TargetSkuId
                Resultado             = "SemAlteracao"
                Detalhe               = "Usuário(a) já possui a licença"
                Executado             = "Verdadeiro"
                Autor                 = $env:USERNAME
                DataHora              = (Get-Date)
            }
            Write-Log -Row $LogRow
            continue
        }

        # Monta o objeto para adicionar APENAS a licença alvo
        $addLicenses = @(@{ SkuId = $TargetSkuId })

        # Atribui apenas a licença alvo, sem remover nenhuma outra
        Set-MgUserLicense -UserId $upn -AddLicenses $addLicenses -RemoveLicenses @() | Out-Null

        Write-Host "`tLicença atribuída 📓"
        $LogRow = [pscustomobject]@{
            UPN                   = $upn
            Acao                  = "Atribuicao"
            Licenca_Alvo          = $TargetSkuId
            Resultado             = "Sucesso"
            Detalhe               = "Licença atribuída"
            Executado             = "Verdadeiro"
            Autor                 = $env:USERNAME
            DataHora              = (Get-Date)
        }
        Write-Log -Row $LogRow

    } catch {
        Write-Warning "Falha ao adicionar licença ✕"
        Write-Host $_.Exception.Message

        $LogRow = [pscustomobject]@{
            UPN                   = $upn
            Acao                  = "Atribuicao"
            Licenca_Alvo          = $TargetSkuId
            Resultado             = "Erro"
            Detalhe               = $_.Exception.Message
            Executado             = "Falso"
            Autor                 = $env:USERNAME
            DataHora              = (Get-Date)
        }
        Write-Log -Row $LogRow
    }
}
