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
if (-not $rows -or -not ($rows[0].PSObject.Properties.Name -contains 'ExchangeGuid')) {
    Write-Error "Input CSV must contain a column named 'ExchangeGuid'."
    exit 1
}

Write-Host "Resolving UPNs for $($rows.Count) Exchange GUIDs..." -ForegroundColor Cyan

# Collect UPN-only output (one per line)
$UPNList = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $guid = ($row.ExchangeGuid).Trim()

    if ([string]::IsNullOrWhiteSpace($guid)) {
        $UPNList.Add("")
        Write-Host ("[{0}/{1}] Blank GUID -> (empty UPN)" -f $index, $rows.Count) `
            -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $guid) `
        -ForegroundColor Yellow

    try {
        $mailbox = $null

        # Direct lookup by Exchange GUID
        $mailbox = Get-Mailbox -Identity $guid -ErrorAction Stop

        if (-not $mailbox) {
            $UPNList.Add("")
            Write-Host "   No mailbox found -> (empty UPN)" -ForegroundColor DarkYellow
            continue
        }

        $UPNList.Add($mailbox.UserPrincipalName)
        Write-Host "   UPN: $($mailbox.UserPrincipalName)" -ForegroundColor Green
    }
    catch {
        $UPNList.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty UPN)"
    }
}

# Write output: single-column CSV
$outputObjects = $UPNList | ForEach-Object {
    [pscustomobject]@{ UserPrincipalName = $_ }
}

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. UPNs written to: $OutputCsv" -ForegroundColor Cyan