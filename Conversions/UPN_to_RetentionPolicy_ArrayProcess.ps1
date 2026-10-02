# Paths
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

Write-Host "Resolving Retention Policies for $($rows.Count) UPNs..." -ForegroundColor Cyan

# Collect RetentionPolicy-only output (one per line)
$RetentionPolicyList = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $upn = ($row.UserPrincipalName).Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        $RetentionPolicyList.Add("")
        Write-Host ("[{0}/{1}] Blank UPN -> (empty RetentionPolicy)" -f $index, $rows.Count) `
            -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $upn) `
        -ForegroundColor Yellow

    try {
        $mailbox = $null

        # Direct lookup by UPN
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop

        if (-not $mailbox) {
            $RetentionPolicyList.Add("")
            Write-Host "   No mailbox found -> (empty RetentionPolicy)" -ForegroundColor DarkYellow
            continue
        }

        $RetentionPolicyList.Add($mailbox.RetentionPolicy)
        Write-Host "   RetentionPolicy: $($mailbox.RetentionPolicy)" -ForegroundColor Green
    }
    catch {
        $RetentionPolicyList.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty RetentionPolicy)"
    }
}

# Write output: single-column CSV
$outputObjects = $RetentionPolicyList | ForEach-Object {
    [pscustomobject]@{ RetentionPolicy = $_ }
}

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. Retention Policies written to: $OutputCsv" -ForegroundColor Cyan