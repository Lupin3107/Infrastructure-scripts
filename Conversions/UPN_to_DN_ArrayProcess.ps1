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
if (-not $rows -or -not ($rows | Get-Member -Name UserPrincipalName -MemberType NoteProperty)) {
    Write-Error "Input CSV must contain a column named 'UserPrincipalName'."
    exit 1
}

Write-Host "Resolving DisplayNames for $($rows.Count) UPNs..." -ForegroundColor Cyan

# Collect DisplayName-only output (one per line)
$displayNameLines = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $upn = ($row.UserPrincipalName).Trim()

    if ([string]::IsNullOrWhiteSpace($upn)) {
        # Keep alignment: blank line for blank input
        $displayNameLines.Add("")
        Write-Host ("[{0}/{1}] Blank UPN -> (empty DisplayName)" -f $index, $rows.Count) -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $upn) -ForegroundColor Yellow

    try {
        # Preferred: direct lookup by UPN (UserId accepts UPN or ObjectId)
        $user = $null
        try {
            $user = Get-MgUser -UserId $upn -ErrorAction Stop
        }
        catch {
            # Fallback: exact filter on userPrincipalName (handles aliases/case issues)
            $exact = Escape-ODataLiteral $upn
            $filter = "userPrincipalName eq '$exact'"
            $user = Get-MgUser -Filter $filter -All -ErrorAction Stop | Select-Object -First 1
        }

        if (-not $user) {
            # No match -> blank line
            $displayNameLines.Add("")
            Write-Host "   No user found -> (empty DisplayName)" -ForegroundColor DarkYellow
            continue
        }

        $displayName = $user.DisplayName
        $displayNameLines.Add($displayName)
        Write-Host "   DisplayName: $displayName" -ForegroundColor Green
    }
    catch {
        # On error, keep alignment with a blank line; log the error to console
        $displayNameLines.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty DisplayName)"
        Start-Sleep -Milliseconds 200
    }
}

# Write output: single column CSV with only DisplayNames
$outputObjects = $displayNameLines | ForEach-Object { [pscustomobject]@{ DisplayName = $_ } }

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Done. DisplayNames written to: $OutputCsv" -ForegroundColor Cyan