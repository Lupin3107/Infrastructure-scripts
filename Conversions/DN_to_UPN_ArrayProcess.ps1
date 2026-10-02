
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
if (-not $rows -or -not ($rows | Get-Member -Name DisplayName -MemberType NoteProperty)) {
    Write-Error "Input CSV must contain a column named 'DisplayName'."
    exit 1
}

Write-Host "Resolving UPNs for $($rows.Count) display names..." -ForegroundColor Cyan

# Collect UPN-only output (one per line)
$upnLines = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $inputName = ($row.DisplayName).Trim()

    if ([string]::IsNullOrWhiteSpace($inputName)) {
        # Keep alignment: blank line for blank input
        $upnLines.Add("")
        Write-Host ("[{0}/{1}] Blank DisplayName -> (empty UPN)" -f $index, $rows.Count) -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $inputName) -ForegroundColor Yellow

    try {
        # 1) Broad search (requires ConsistencyLevel eventual)
        $searchQuery = "displayName:$inputName"
        $users = Get-MgUser -Search $searchQuery -ConsistencyLevel eventual -All -ErrorAction Stop

        # 2) Fallback to exact filter if none
        if (-not $users -or $users.Count -eq 0) {
            $exact = Escape-ODataLiteral $inputName
            $filter = "displayName eq '$exact'"
            $users = Get-MgUser -Filter $filter -All -ErrorAction Stop
        }

        if (-not $users -or $users.Count -eq 0) {
            # No match -> blank line
            $upnLines.Add("")
            Write-Host "   No user found" -ForegroundColor DarkYellow
            continue
        }

        # Strategy: prefer exact displayName matches, else take first match.
        $exactMatches = $users | Where-Object {
            $_.DisplayName -eq $inputName -or ($_.DisplayName -as [string]).ToLower() -eq $inputName.ToLower()
        }

        $selected = if ($exactMatches -and $exactMatches.Count -gt 0) {
            $exactMatches | Select-Object -First 1
        } else {
            $users | Select-Object -First 1
        }

        $upn = $selected.UserPrincipalName
        $upnLines.Add($upn)
        Write-Host "   UPN: $upn" -ForegroundColor Green
    }
    catch {
        # On error, keep alignment with a blank line; log the error to console
        $upnLines.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty UPN)"
        Start-Sleep -Milliseconds 200
    }
}

# Write output: single column CSV with only UPNs
# Construct a simple array of PSCustomObjects with one property 'UserPrincipalName'
$outputObjects = $upnLines | ForEach-Object { [pscustomobject]@{ UserPrincipalName = $_ } }

# Ensure output directory exists
$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Done. UPNs written to: $OutputCsv" -ForegroundColor Cyan