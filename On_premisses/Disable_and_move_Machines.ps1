# ======================
# Parâmetros e setup
# ======================
$HostsList = ".\Example.txt"
$LogFolder = "C:\Logs\Path\Example"
$Encoding  = "utf8"
if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }

$TimeTag = Get-Date -Format 'yyyyMMdd_HHmm'
$LogFile = Join-Path $LogFolder "Machines_Disables_$TimeTag.csv"

function Log {
    param(
        [string]$Hostname, [string]$Status, [string]$Msg,
        [bool]$Exec = $false, [string]$OU, [string]$DN,
        [string]$Err, [string]$Act
    )
    $row = [pscustomobject]@{
        Host        = $Hostname
        Status      = $Status
        Mensagem    = $Msg
        Executado   = ($(if ($Exec) { 'Verdadeiro' } else { 'Falso' }))
        Autor       = $env:USERNAME
        Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        OUAtual     = $OU
        DN          = $DN
        ErroDetalhe = $Err
        Acao        = $Act
    }
    $exists = Test-Path -LiteralPath $LogFile
    $row | Export-Csv -Path $LogFile -Encoding $Encoding -NoTypeInformation -Append:$exists
}

# ======================
# Parâmetros de AD
# ======================
$Description = "Disable: [$TimeTag] Obsoleto/Rotina."
$TargetOU    = "OU=DISABLE,OU=OBJECTS,OU=DOMAIN.COM.BR,DC=domain,DC=com,DC=br"
$AllowedOUs  = @(
    "OU=ALLOWED_ONE,DC=DOMAIN,DC=com,DC=br",
    "OU=ALLOWED_TWO,DC=DOMAIN,DC=com,DC=br",
    "OU=ALLOWED_THREE,DC=DOMAIN,DC=com,DC=br"
)

# ======================
# Input
# ======================
if (-not (Test-Path $HostsList)) { Write-Error "Arquivo não encontrado: $HostsList"; exit }
$Suspects = Get-Content $HostsList | ForEach-Object { $_.Trim() } | Where-Object { $_ }
$Total    = $Suspects.Count

Write-Host "Validando $Total hosts (ping único às $((Get-Date).ToString('HH:mm')))..." -ForegroundColor Cyan
Write-Host "Regra: ONLINE=salvo | OFFLINE=desabilita+move"

# ======================
# Verificação de ping
# ======================
$Offline = @()
foreach ($h in $Suspects) {
    $c = $h.Trim()
    $ok = Test-Connection -ComputerName $c -Count 2 -Quiet
    if ($ok) {
        Write-Host "$c [ONLINE]" -ForegroundColor Green
        Log -Host $c -Status "Online" -Msg "Respondeu ao ping; sem ação." -Exec:$false -Act "Nenhuma"
    } else {
        Write-Host "$c [OFFLINE]" -ForegroundColor Red
        $Offline += $c
        Log -Host $c -Status "Offline" -Msg "Sem resposta ao ping; elegível." -Exec:$false
    }
}

if (-not $Offline) { Write-Host "`nTodos responderam ao ping. Nada a fazer." -ForegroundColor Green; exit }
$Suspects = $Offline

# ======================
# Processamento no AD
# ======================
foreach ($h in $Suspects) {
    $ad = Get-ADComputer -Filter "Name -like '$h'" -Properties DistinguishedName,Enabled,Description -ErrorAction SilentlyContinue
    if (-not $ad) {
        Write-Host "$h [NÃO ENCONTRADO]" -ForegroundColor Yellow
        Log $h "Não encontrado" "Objeto não localizado no AD." $false $null $null $null "Nenhuma"
        continue
    }

    $dn = $ad.DistinguishedName
    if (-not ($AllowedOUs | Where-Object { $dn -like "*$_" })) {
        Write-Host "$h [FORA DE OU PERMITIDA]" -ForegroundColor Yellow
        Log $h "Fora da OU permitida" "Fora das OUs autorizadas." $false $null $dn $null "Nenhuma"
        continue
    }

    # Já desativado?
    if (-not $ad.Enabled) {
        $acts=@(); $eDesc=$null; $eMove=$null

        if ($dn -like "*$TargetOU") {
            # Only set description if empty
            if ([string]::IsNullOrWhiteSpace($ad.Description)) {
                try { Set-ADComputer -Identity $ad.Name -Description $Description -ErrorAction Stop; $acts += "SetDescription" }
                catch { $eDesc = $_.Exception.ToString(); Write-Host "$h [ERRO SetDescription]"}
            }
        } else {
            # Set description if empty and move
            if ([string]::IsNullOrWhiteSpace($ad.Description)) {
                try { Set-ADComputer -Identity $ad.Name -Description $Description -ErrorAction Stop; $acts += "SetDescription" }
                catch { $eDesc = $_.Exception.ToString(); Write-Host "$h [ERRO SetDescription]"}
            }

            try {
                Move-ADObject -Identity $dn -TargetPath $TargetOU -ErrorAction Stop; $acts += "Move"
            }
            catch { $eMove = $_.Exception.ToString(); Write-Host "$h [ERRO Move]"}
        }

        if ($acts.Count) {
            $a = $acts -join "+"
            Write-Host "$h [JÁ DESATIVADO/$a]" -ForegroundColor Yellow
            Log $h "Já desativado" "Ações: $a." $true $TargetOU $dn $null $a
        } else {
            Write-Host "$h [JÁ DESATIVADO]" -ForegroundColor Yellow
            Log $h "Já desativado" "Sem alterações." $false $TargetOU $dn $null "Nenhuma"
        }

        if ($eDesc) { Log $h "Erro" "Falha ao definir descrição (desativado)." $false $null $dn $eDesc "Falha SetDescription" }
        if ($eMove) { Log $h "Erro" "Falha ao mover para OU alvo." $false $null $dn $eMove "Falha Move" }
        continue
    }

    # Desativar + mover + descrever
    try {
        Disable-ADAccount -Identity $dn -ErrorAction Stop
        Move-ADObject    -Identity $dn -TargetPath $TargetOU -ErrorAction Stop
        Set-ADComputer   -Identity $ad.Name -Description $Description -ErrorAction Stop

        Write-Host "$h [DESABILITADO+MOVIDO]" -ForegroundColor Red
        Log $h "Desabilitado" "Desabilitado, movido e descrição atualizada." $true $TargetOU $dn $null "Disable+Move+SetDescription"
    }
    catch {
        Write-Host "$h [ERRO Disable/Move/SetDescription] $($_.Exception.Message)"
        Log $h "Erro" "Falha ao desabilitar/mover/atualizar descrição." $false $null $dn $_.Exception.ToString() "Falha Disable/Move/SetDescription"
    }
}