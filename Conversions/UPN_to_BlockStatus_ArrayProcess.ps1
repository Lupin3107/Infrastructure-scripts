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

Write-Host "Checking account block status for $($rows.Count) UPNs..." -ForegroundColor Cyan

# Collect BlockStatus-only output (one per line)
$BlockStatusList = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $upn = ($row.UserPrincipalName).Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        $BlockStatusList.Add("")
        Write-Host ("[{0}/{1}] Blank UPN -> (empty status)" -f $index, $rows.Count) `
            -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Checking: {2}" -f $index, $rows.Count, $upn) `
        -ForegroundColor Yellow

    try {
        $user = $null

        # Direct lookup by UPN
        $user = Get-MgUser `
            -UserId $upn `
            -Property AccountEnabled `
            -ErrorAction Stop

        if (-not $user) {
            $BlockStatusList.Add("")
            Write-Host "   User not found -> (empty status)" -ForegroundColor DarkYellow
            continue
        }

        if ($user.AccountEnabled -eq $false) {
            $BlockStatusList.Add("Yes")
            Write-Host "   Blocked: Yes" -ForegroundColor Red
        }
        else {
            $BlockStatusList.Add("No")
            Write-Host "   Blocked: No" -ForegroundColor Green
        }
    }
    catch {
        $BlockStatusList.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty status)"
    }
}

# Write output: single-column CSV
$outputObjects = $BlockStatusList | ForEach-Object {
    [pscustomobject]@{ Blocked = $_ }
}

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. Block status written to: $OutputCsv" -ForegroundColor Cyan