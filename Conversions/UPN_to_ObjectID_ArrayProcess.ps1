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

Write-Host "Resolving Object IDs for $($rows.Count) UPNs..." -ForegroundColor Cyan

# Collect ObjectId-only output
$ObjectIdList = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $upn = ($row.UserPrincipalName).Trim()

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $upn) `
        -ForegroundColor Yellow

    try {
        # Lookup in Entra ID via Graph
        $user = Get-MgUser -UserId $upn -Property Id -ErrorAction Stop

        if (-not $user) {
            $ObjectIdList.Add("")
            Write-Host "   User not found -> (empty ObjectId)" -ForegroundColor DarkYellow
            continue
        }

        $ObjectIdList.Add($user.Id)
        Write-Host "   ObjectId: $($user.Id)" -ForegroundColor Green
    }
    catch {
        $ObjectIdList.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty ObjectId)"
    }
}

# Write output: single-column CSV
$outputObjects = $ObjectIdList | ForEach-Object {
    [pscustomobject]@{ ObjectId = $_ }
}

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. Object IDs written to: $OutputCsv" -ForegroundColor Cyan