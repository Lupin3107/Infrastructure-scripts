# Set the DL display name or SMTP address ---
$GroupIdentity = "Example@domain.com.br"

# --- Ensure output directory exists and define output path ---
$OutPath = "C:\Example\path\DL_members.csv"
$OutDir  = Split-Path $OutPath -Parent
if (-not (Test-Path -Path $OutDir)) {
    New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
}

# Get members (handles Users, Shared Mailboxes, Mail Contacts, etc.) ---
try {
    $members = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
} catch {
    Write-Error "Error retrieving members for '$GroupIdentity'. $_"
    exit 1
}

# Export members (handle scalar vs array and include useful fields) ---
if ($members) {
    # Force array to avoid scalar behavior
    $members = @($members)

    $members |
        Select-Object `
            @{Name='Group';               Expression = { $GroupIdentity }}, `
            @{Name='Name';                Expression = { $_.Name }}, `
            @{Name='DisplayName';         Expression = { $_.DisplayName }}, `
            @{Name='PrimarySmtpAddress';  Expression = { $_.PrimarySmtpAddress }} |
        Export-Csv -Path $OutPath -NoTypeInformation -Encoding UTF8 -Force

    Write-Host "Export complete: $OutPath"
} else {
    Write-Host "No members returned for $GroupIdentity"
}
