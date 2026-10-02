# Paths (static as requested)
$InputCsv  = "c:\Example\path\input.csv"
$OutputCsv = "C:\Example\path\result.csv"

# Ensure input exists
if (-not (Test-Path -LiteralPath $InputCsv)) {
    Write-Error "Input CSV not found: $InputCsv"
    exit 1
}

# Import and validate column
$rows = Import-Csv -LiteralPath $InputCsv
if (-not $rows -or -not ($rows[0].PSObject.Properties.Name -contains 'UserPrincipalName')) {
    Write-Error "Input CSV must contain a column named 'UserPrincipalName'."
    exit 1
}

Write-Host "Resolving MailNicknames for $($rows.Count) UPNs..." -ForegroundColor Cyan

# Collect MailNickname-only output (one per line)
$Mailnicknamelist = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $upn = ($row.UserPrincipalName).Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        $Mailnicknamelist.Add("")
        Write-Host ("[{0}/{1}] Blank UPN -> (empty MailNickname)" -f $index, $rows.Count) -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $upn) -ForegroundColor Yellow

    try {
        $user = $null

        # Preferred: direct lookup by UPN (fastest)
        try {
            $user = Get-MgUser -UserId $upn -Property MailNickname -ErrorAction Stop
        }
        catch {
            # Even simpler fallback: exact filter, no escaping helpers
            $user = Get-MgUser `
                -Filter "userPrincipalName eq '$upn'" `
                -Property MailNickname `
                -All `
                -ErrorAction Stop |
                Select-Object -First 1
        }

        if (-not $user) {
            $Mailnicknamelist.Add("")
            Write-Host "   No user found -> (empty MailNickname)" -ForegroundColor DarkYellow
            continue
        }

        $Mailnicknamelist.Add($user.MailNickname)
        Write-Host "   MailNickname: $($user.MailNickname)" -ForegroundColor Green
    }
    catch {
        $Mailnicknamelist.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty MailNickname)"
    }
}

# Write output: single-column CSV
$outputObjects = $Mailnicknamelist | ForEach-Object {
    [pscustomobject]@{ MailNickname = $_ }
}

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. MailNicknames written to: $OutputCsv" -ForegroundColor Cyan