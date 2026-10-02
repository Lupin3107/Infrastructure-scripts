$csvPath = "C:\Example\path\Litigation_hold_results.csv"
$logFolder = "C:\Example"
$logFile = Join-Path $logFolder "Litigation_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Create log folder if it doesn't exist
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# Load users from CSV
$users = Import-Csv -Path $csvPath -Encoding UTF8
$logs = @()

foreach ($user in $users) {
    $upn = $user.UPN.Trim()
    if (-not $upn) {
        Write-Host "! UPN inválido ou ausente." -ForegroundColor DarkYellow
        continue
    }

    $logEntry = [PSCustomObject]@{
        UPN       = $upn
        Status    = ""
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop

        if ($mailbox.LitigationHoldEnabled) {
            Write-Host "✓ Litigation Hold já está habilitado para $upn." -ForegroundColor Green
            $logEntry.Status = "Already Enabled"
        } else {
            Set-Mailbox -Identity $upn -LitigationHoldEnabled $true -ErrorAction Stop
            Write-Host "✕ Litigation Hold habilitado para $upn." -ForegroundColor Magenta
            $logEntry.Status = "Enabled"
        }
    } catch {
        $errorMessage = $_.Exception.Message
        $logEntry | Add-Member -MemberType NoteProperty -Name ErrorMessage -Value $errorMessage

        if ($errorMessage -like "*licença*" -or $errorMessage -like "*license*") {
            Write-Host "! Problema de licença para $upn." -ForegroundColor Red
            $logEntry.Status = "License Restriction"
        } else {
            Write-Host "↻ Mailbox não encontrado para $upn." -ForegroundColor Yellow
            $logEntry.Status = "Mailbox Not Found"
        }
    }

    $logs += $logEntry
}

# Export logs to CSV
$logs | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
Write-Host "✓ Log salvo em: $logFile" -ForegroundColor Cyan