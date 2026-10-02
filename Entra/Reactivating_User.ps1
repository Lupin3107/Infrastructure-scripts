
<# Ações em massa (AD On-Premises):
    - Ativar a conta (Enable-ADAccount)
    - Desmarcar "User must change password at next logon"
      -> Set-ADUser -ChangePasswordAtLogon:$false
      -> Ajustar pwdLastSet = -1 (garante limpeza)
    - Mover o usuário para a OU alvo via Move-ADObject

    Estrutura com CSV, logs, validações e tratamento de erros.
#>

Import-Module ActiveDirectory -ErrorAction Stop

# --- Paths ---
$csvPath   = "C:\Example\path\DisableUsers.csv"   # CSV com 'UserPrincipalName' (ou 'SamAccountName')
$logFolder = "C:\Example\path"
$logFile   = Join-Path $logFolder "OnPrem_Enable_Clear_Move_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# --- OU de destino ---
$targetOUPath = "OU=Example,OU=path,DC=domain,DC=com,DC=br"

# --- Preparar pasta de log ---
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# --- Carregar usuários do CSV ---
$users = Import-Csv -Path $csvPath
$logs  = @()

foreach ($row in $users) {
    # Aceita UPN (preferência) ou sAMAccountName
    $upn = $row.UserPrincipalName
    $sam = $row.SamAccountName

    # Entrada base de log
    $logEntry = [PSCustomObject]@{
        UserPrincipalName = $upn
        SamAccountName    = $sam
        Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Status            = ""
        Details           = ""
        Enabled           = $null
        ClearedChangeAtLogon = $null
        MovedToOU         = $false
    }

    try {
        # --- Localizar usuário de forma robusta ---
        $user = $null
        if ($upn) {
            $user = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties Enabled,PasswordLastSet,pwdLastSet -ErrorAction Stop
        }
        if (-not $user -and $sam) {
            $user = Get-ADUser -Identity $sam -Properties Enabled,PasswordLastSet,pwdLastSet -ErrorAction Stop
        }
        if (-not $user) { throw [System.Exception] "User not found by UPN or SamAccountName." }

        # --- 1) Ativar conta (Enable-ADAccount) ---
        if ($user.Enabled -eq $true) {
            Write-Host "Conta JÁ ATIVA para $($user.SamAccountName)." -ForegroundColor Green
            $logEntry.Enabled = "Already Enabled"
        } else {
            Enable-ADAccount -Identity $user.DistinguishedName  # [1](https://learn.microsoft.com/en-us/powershell/module/activedirectory/enable-adaccount?view=windowsserver2025-ps)
            # Recarregar estado
            $user = Get-ADUser -Identity $user.DistinguishedName -Properties Enabled,PasswordLastSet,pwdLastSet
            if ($user.Enabled -eq $true) {
                Write-Host "Conta ATIVADA para $($user.SamAccountName)." -ForegroundColor Magenta
                $logEntry.Enabled = "Enabled"
            } else {
                Write-Host "Tentativa de ativar sem confirmação para $($user.SamAccountName)." -ForegroundColor Yellow
                $logEntry.Enabled = "Enable Attempted – Not Confirmed"
                $logEntry.Details += " Verifique permissões/replicação;"
            }
        }

        # --- 2) Desmarcar "User must change password at next logon" ---
        # a) Limpar flag via parâmetro (ChangePasswordAtLogon)
        Set-ADUser -Identity $user.DistinguishedName -ChangePasswordAtLogon:$false  # [2](https://learn.microsoft.com/en-us/powershell/module/activedirectory/set-aduser?view=windowsserver2025-ps)
        # b) Garantir limpeza do estado forçado: pwdLastSet = -1
        Set-ADUser -Identity $user.DistinguishedName -Replace @{pwdLastSet = -1}     # [3](https://learn.microsoft.com/en-us/windows/win32/adsi/user-must-change-password-at-next-logon)

        # Checagem simples: se pwdLastSet é diferente de 0, não há exigência de troca no próximo logon
        $check = Get-ADUser -Identity $user.DistinguishedName -Properties pwdLastSet
        if ($check.pwdLastSet -ne 0) {
            Write-Host "Flag 'Change password at next logon' LIMPA para $($user.SamAccountName)." -ForegroundColor Cyan
            $logEntry.ClearedChangeAtLogon = $true
        } else {
            Write-Host "Flag de 'change at next logon' ainda presente para $($user.SamAccountName)." -ForegroundColor Yellow
            $logEntry.ClearedChangeAtLogon = $false
            $logEntry.Details += " pwdLastSet ainda 0;"
        }

        # --- 3) Mover o usuário para a OU de destino ---
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $targetOUPath  # [4](https://learn.microsoft.com/en-us/powershell/module/activedirectory/move-adobject?view=windowsserver2025-ps)

        # Confirmar que DN agora está sob a OU alvo
        $moved = Get-ADUser -Filter "SamAccountName -eq '$($user.SamAccountName)'" -Properties DistinguishedName
        if ($moved.DistinguishedName -like "*OU=Israel,OU=Futuragene,OU=Futuragene.corp,OU=SuzanoBahiaSul,OU=Objetos,DC=suzano,DC=com,DC=br") {
            Write-Host "Usuário movido para OU de destino: $($moved.DistinguishedName)" -ForegroundColor Green
            $logEntry.MovedToOU = $true
            $logEntry.Status    = "Enabled + Cleared + Moved"
        } else {
            Write-Host "Movimentação não confirmada para $($user.SamAccountName)." -ForegroundColor Yellow
            $logEntry.MovedToOU = $false
            $logEntry.Status    = "Enabled + Cleared + Move Attempted – Not Confirmed"
        }

    } catch {
        $errorMessage = $_.Exception.Message

        switch ($true) {
            # Permissões insuficientes / autorização negada
            { $errorMessage -like "*Access is denied*" -or
              $errorMessage -like "*insufficient privileges*" } {
                Write-Host "Permissão insuficiente para $upn / $sam." -ForegroundColor Red
                $logEntry.Status  = "Authorization Denied"
                $logEntry.Details = $errorMessage
                break
            }

            # Usuário não encontrado
            { $errorMessage -like "*User not found*" -or
              $errorMessage -like "*cannot be found*" -or
              $errorMessage -like "*not found by UPN or sAMAccountName*" } {
                Write-Host "Usuário não encontrado: $upn / $sam." -ForegroundColor Yellow
                $logEntry.Status  = "User Not Found"
                $logEntry.Details = $errorMessage
                break
            }

            # Erros ao mover (ex.: fora do forest, OU protegida, RID Master, etc.)
            { $errorMessage -like "*The requested operation could not be performed*" -or
              $errorMessage -like "*protected from accidental deletion*" -or
              $errorMessage -like "*naming context could not be found*" } {
                Write-Host "Falha ao mover para a OU destino ($($targetOUPath)) para $upn / $sam." -ForegroundColor Yellow
                $logEntry.Status  = "Move Error"
                $logEntry.Details = $errorMessage
                break
            }

            default {
                Write-Host "Erro inesperado para $upn / $sam $errorMessage" -ForegroundColor Red
                $logEntry.Status  = "Error"
                $logEntry.Details = $errorMessage
            }
        }
    }

    $logs += $logEntry
}

# --- Exportar log ---
$logs | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8BOM
Write-Host "✅ Log salvo em $logFile" -ForegroundColor Cyan
 