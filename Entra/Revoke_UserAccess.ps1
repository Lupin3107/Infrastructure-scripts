
<# Revoga as sessões de login (refresh tokens) de usuários em massa.
    - Gera log por usuário com status e detalhes
#>

# --- Paths ---
$csvPath   = "C:\Example\path\DisableUsers.csv"   # CSV com coluna 'UserPrincipalName'
$logFolder = "C:\Example\path"
$logFile   = Join-Path $logFolder "Revoke_Sessions_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# --- Preparar pasta de log ---
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# --- Carregar usuários do CSV ---
$users = Import-Csv -Path $csvPath
$logs  = @()

foreach ($row in $users) {
    $upn = $row.UserPrincipalName

    # Entrada base de log
    $logEntry = [PSCustomObject]@{
        UserPrincipalName = $upn
        Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Status            = ""
        Details           = ""
        RevokedSessions   = $false
    }

    try {
        # Obter o usuário (UPN aceito diretamente)
        $user = Get-MgUser -UserId $upn -ErrorAction Stop

        # Revogar sessões (invalidate refresh tokens)
        Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop

        Write-Host "Sessões REVOGADAS para $upn." -ForegroundColor Cyan
        $logEntry.Status          = "Sessions Revoked"
        $logEntry.Details         = "Refresh tokens invalidados; usuário precisará reautenticar."
        $logEntry.RevokedSessions = $true

    } catch {
        $errorMessage = $_.Exception.Message

        switch ($true) {
            # Permissão insuficiente / autorização negada
            { $errorMessage -like "*Insufficient privileges*" -or
              $errorMessage -like "*Authorization_RequestDenied*" -or
              $errorMessage -like "*does not have permission*" } {
                Write-Host "Permissão insuficiente para $upn." -ForegroundColor Red
                $logEntry.Status  = "Authorization Denied"
                $logEntry.Details = $errorMessage
                break
            }

            # Usuário não encontrado
            { $errorMessage -like "*Resource*not found*" -or
              $errorMessage -like "*cannot be found*" -or
              $errorMessage -like "*User not found*" } {
                Write-Host "Usuário $upn não encontrado." -ForegroundColor Yellow
                $logEntry.Status  = "User Not Found"
                $logEntry.Details = $errorMessage
                break
            }

            default {
                Write-Host "Erro inesperado para $upn $errorMessage" -ForegroundColor Red
                $logEntry.Status  = "Error"
                $logEntry.Details = $errorMessage
            }
        }
    }

    $logs += $logEntry
}

# --- Exportar log para CSV ---
$logs | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
Write-Host "✅ Log salvo em $logFile" -ForegroundColor Cyan