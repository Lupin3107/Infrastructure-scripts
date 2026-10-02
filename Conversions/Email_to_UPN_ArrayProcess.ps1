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
if (-not $rows -or -not ($rows | Get-Member -Name EmailAddress -MemberType NoteProperty)) {
    Write-Error "Input CSV must contain a column named 'EmailAddress'."
    exit 1
}

Write-Host "Resolving UPNs for $($rows.Count) email addresses..." -ForegroundColor Cyan

# Collect UPN-only output
$upnLines = New-Object System.Collections.Generic.List[string]

$index = 0
foreach ($row in $rows) {
    $index++
    $email = ($row.EmailAddress).Trim()

    if ([string]::IsNullOrWhiteSpace($email)) {
        $upnLines.Add("")
        Write-Host ("[{0}/{1}] Blank email -> (empty UPN)" -f $index, $rows.Count) -ForegroundColor DarkYellow
        continue
    }

    Write-Host ("[{0}/{1}] Resolving: {2}" -f $index, $rows.Count, $email) -ForegroundColor Yellow

    try {
        $user = Get-MgUser -Filter "mail eq '$email'" -ConsistencyLevel eventual -All -ErrorAction Stop |
                Select-Object -First 1

        if (-not $user) {
            $user = Get-MgUser -Filter "userPrincipalName eq '$email'" -ConsistencyLevel eventual -All -ErrorAction Stop |
                    Select-Object -First 1
        }

        if (-not $user) {
            $upnLines.Add("")
            Write-Host "   No user found -> (empty UPN)" -ForegroundColor DarkYellow
            continue
        }

        $upnLines.Add($user.UserPrincipalName)
        Write-Host "   UPN: $($user.UserPrincipalName)" -ForegroundColor Green
    }
    catch {
        $upnLines.Add("")
        Write-Warning "   Error: $($_.Exception.Message) -> (empty UPN)"
    }
}

# Single-column output CSV
$outputObjects = $upnLines | ForEach-Object {
    [pscustomobject]@{
        UserPrincipalName = $_
    }
}

$dir = Split-Path -Path $OutputCsv
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$outputObjects | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. UPNs written to: $OutputCsv" -ForegroundColor Cyan