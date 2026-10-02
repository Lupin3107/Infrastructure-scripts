# Paths
$csvPath  = "C:\Example\path\Archive_enable.csv"
$logFolder = "C:\Example\path"
$logFile   = Join-Path $logFolder "Archive_Enable_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Create log folder if it doesn't exist
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory | Out-Null
}

# Load users from CSV (expects a column named 'UserPrincipalName')
$users = Import-Csv -Path $csvPath
$logs  = @()

foreach ($user in $users) {
    $upn = $user.UserPrincipalName

    # Prepare log entry
    $logEntry = [PSCustomObject]@{
        UserPrincipalName = $upn
        Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    try {
        # Get mailbox
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop

        # Determine archive status robustly
        $archiveEnabled = $false
        if ($mailbox.PSObject.Properties.Name -contains 'ArchiveStatus') {
            $archiveEnabled = ($mailbox.ArchiveStatus -eq 'Active')
        } elseif ($mailbox.PSObject.Properties.Name -contains 'ArchiveGuid') {
            $archiveEnabled = ($null -ne $mailbox.ArchiveGuid)
        }

        if ($archiveEnabled) {
            Write-Host "Archive is ALREADY ENABLED for $upn." -ForegroundColor Green
            $logEntry.Status  = "Already Enabled"
            $logEntry.Details = "ArchiveStatus: $($mailbox.ArchiveStatus)"
        } else {
            # Enable archive
            Enable-Mailbox -Identity $upn -Archive -ErrorAction Stop

            # Re-check to confirm
            $mailboxAfter = Get-Mailbox -Identity $upn -ErrorAction Stop
            $archiveNow = $false
            if ($mailboxAfter.PSObject.Properties.Name -contains 'ArchiveStatus') {
                $archiveNow = ($mailboxAfter.ArchiveStatus -eq 'Active')
            } elseif ($mailboxAfter.PSObject.Properties.Name -contains 'ArchiveGuid') {
                $archiveNow = ($null -ne $mailbox.ArchiveGuid)
            }

            if ($archiveNow) {
                Write-Host "Archive has been ENABLED for $upn." -ForegroundColor Magenta
                $logEntry.Status  = "Enabled"
                $logEntry.Details = "ArchiveStatus: $($mailboxAfter.ArchiveStatus)"
            } else {
                Write-Host "Archive enable attempted but status not confirmed for $upn." -ForegroundColor Yellow
                $logEntry.Status  = "Enable Attempted – Not Confirmed"
                $logEntry.Details = "Check licensing/policy; ArchiveStatus=$($mailboxAfter.ArchiveStatus)"
            }
        }
    } catch {
        $errorMessage = $_.Exception.Message

        switch ($true) {
            # Licensing restrictions (Portuguese sample text and typical English messages)
            { $errorMessage -like "*licença*" -or
              $errorMessage -like "*A licença do Microsoft Exchange Online não permite*" -or
              $errorMessage -like "*The user does not have a valid license*" } {
                Write-Host "Licensing issue for $upn. Cannot enable Archive." -ForegroundColor Red
                $logEntry.Status  = "License Restriction"
                $logEntry.Details = $errorMessage
                break
            }

            # Mailbox not found / identity errors
            { $errorMessage -like "*cannot be found*" -or
              $errorMessage -like "*No corresponding mailbox*" -or
              $errorMessage -like "*User not found*" } {
                Write-Host "User $upn not found or does not have a mailbox." -ForegroundColor Yellow
                $logEntry.Status  = "Mailbox Not Found"
                $logEntry.Details = $errorMessage
                break
            }

            # Already has in-place archive errors (rare race condition)
            { $errorMessage -like "*Archive mailbox already exists*" } {
                Write-Host "Archive already exists for $upn." -ForegroundColor Green
                $logEntry.Status  = "Already Enabled (Detected by Error)"
                $logEntry.Details = $errorMessage
                break
            }

            default {
                Write-Host "Unexpected error for $upn $errorMessage" -ForegroundColor Red
                $logEntry.Status  = "Error"
                $logEntry.Details = $errorMessage
            }
        }
    }

    $logs += $logEntry
}

# Export logs to CSV
$logs | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
Write-Host "Log file saved to $logFile ✓" -ForegroundColor Cyan