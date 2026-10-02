<#
.SYNOPSIS
Exports all user mailboxes assigned to a specified Retention Policy (Default: 'Default MRM Policy').

.DESCRIPTION
This script retrieves all user mailboxes with the given RetentionPolicy and exports a CSV with basic mailbox details.

.PARAMETER PolicyName
The name of the Retention Policy to filter on. Defaults to 'Default MRM Policy'.

.PARAMETER OutputPath
Full path to the CSV file to create. Defaults to 'C:\Logs\MRMPolicy_Default_mailboxes.csv'.

#>

param(
    [Parameter(Mandatory=$false)]
    [string]$PolicyName = 'Default MRM Policy',

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = 'C:\Logs\Exchange\MRMPolicy_Default_mailboxes.csv'
)

# Ensure Get-Mailbox is available
if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Write-Error "Get-Mailbox cmdlet not found. Run this script from the Exchange Management Shell or connect to Exchange Online (Import-Module ExchangeOnlineManagement and Connect-ExchangeOnline)."
    exit 1
}

try {
    # Ensure output directory exists
    $outDir = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    Write-Host "Retrieving user mailboxes with RetentionPolicy='$PolicyName'..."

    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
        $_.RecipientTypeDetails -eq 'UserMailbox' -and $_.RetentionPolicy -ne $null -and $_.RetentionPolicy -eq $PolicyName
    }

    # Filter only active accounts
    $activeMailboxes = $mailboxes | Where-Object { $_.AccountDisabled -ne $true }

    $exportData = $activeMailboxes | Select-Object DisplayName, UserPrincipalName, RetentionPolicy, PrimarySmtpAddress

    if ($exportData -and $exportData.Count -gt 0) {
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Force
        Write-Host "Export completed. File saved to $OutputPath"
    }
    else {
        Write-Host "No active user mailboxes found with RetentionPolicy='$PolicyName'. No file created."
    }
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    exit 1
}
