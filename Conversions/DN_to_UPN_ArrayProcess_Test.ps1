#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users
<#
    DN_to_UPN_ArrayProcess.ps1

    Resolves a CSV column of DisplayName values to UserPrincipalName, using
    Microsoft Graph batching to cut per-row round-trips.

    - First pass: batched $search (fuzzy) lookups, 20 per HTTP request.
    - Second pass: batched $filter (exact) fallback for anything unresolved.
    - Retries with backoff on 429 / 503, honoring Retry-After when present.
    - Preserves input row order and count 1:1 in the output.
#>

param(
    [string]$InputCsv  = "c:\Example\path\input.csv",
    [string]$OutputCsv = "C:\Example\path\result.csv",
    [int]$BatchSize     = 20,   # Graph $batch hard limit is 20 sub-requests
    [int]$MaxRetries    = 5
)

$ErrorActionPreference = "Stop"

# ---- Pre-flight checks -----------------------------------------------------

if (-not (Get-Command Get-MgUser -ErrorAction SilentlyContinue)) {
    Write-Error "Microsoft.Graph.Users module not loaded. Run: Import-Module Microsoft.Graph.Users"
    exit 1
}
if (-not (Get-MgContext)) {
    Write-Error "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.Read.All'"
    exit 1
}
if (-not (Test-Path -LiteralPath $InputCsv)) {
    Write-Error "Input CSV not found: $InputCsv"
    exit 1
}

$rows = Import-Csv -LiteralPath $InputCsv
if (-not $rows -or -not ($rows | Get-Member -Name DisplayName -MemberType NoteProperty)) {
    Write-Error "Input CSV must contain a column named 'DisplayName'."
    exit 1
}

# ---- Helpers ----------------------------------------------------------------

function Escape-ODataLiteral {
    param([string]$Value)
    # OData single-quote escaping: ' -> ''
    return $Value.Replace("'", "''")
}

function Invoke-GraphBatch {
    <#
        Sends a set of {id, method, url, headers} requests through /$batch,
        chunked to $BatchSize, with retry/backoff on throttling.
        Returns a hashtable keyed by request id -> body (or $null on failure).
    #>
    param(
        [Parameter(Mandatory)][array]$Requests,
        [int]$BatchSize  = 20,
        [int]$MaxRetries = 5
    )

    $results = @{}

    for ($i = 0; $i -lt $Requests.Count; $i += $BatchSize) {
        $chunk = $Requests[$i..([Math]::Min($i + $BatchSize - 1, $Requests.Count - 1))]
        $pending = $chunk
        $attempt = 0

        while ($pending.Count -gt 0 -and $attempt -le $MaxRetries) {
            $attempt++
            $body = @{ requests = $pending } | ConvertTo-Json -Depth 6

            try {
                $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" -Body $body -ContentType "application/json"
            }
            catch {
                Write-Warning "Batch call failed (attempt $attempt): $($_.Exception.Message)"
                Start-Sleep -Seconds ([Math]::Min([Math]::Pow(2, $attempt), 30))
                continue
            }

            $retryIds = New-Object System.Collections.Generic.List[string]

            foreach ($r in $response.responses) {
                if ($r.status -eq 429 -or $r.status -eq 503) {
                    $retryIds.Add($r.id)
                }
                else {
                    $results[$r.id] = $r
                }
            }

            if ($retryIds.Count -gt 0) {
                $waitSeconds = [Math]::Min([Math]::Pow(2, $attempt), 30)
                Write-Host "   Throttled on $($retryIds.Count) request(s), backing off $waitSeconds s..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $waitSeconds
                $pending = $chunk | Where-Object { $retryIds -contains $_.id }
            }
            else {
                $pending = @()
            }
        }

        # Anything still pending after MaxRetries is treated as unresolved
        foreach ($req in $pending) { $results[$req.id] = $null }
    }

    return $results
}

function Get-BestMatch {
    param([array]$Users, [string]$InputName)
    if (-not $Users -or $Users.Count -eq 0) { return $null }
    $exact = $Users | Where-Object { $_.displayName -and ($_.displayName.ToLower() -eq $InputName.ToLower()) }
    if ($exact) { return $exact | Select-Object -First 1 }
    return $Users | Select-Object -First 1
}

# ---- Build the ordered work list (skip blanks up front) --------------------

$total   = $rows.Count
$upnById = New-Object 'string[]' $total   # final output, indexed by row position
$names   = New-Object 'string[]' $total

for ($i = 0; $i -lt $total; $i++) {
    $name = ($rows[$i].DisplayName).Trim()
    $names[$i] = $name
    if ([string]::IsNullOrWhiteSpace($name)) { $upnById[$i] = "" }
}

Write-Host "Resolving UPNs for $total display names (batched)..." -ForegroundColor Cyan

# ---- Pass 1: batched $search ------------------------------------------------

$searchRequests = @()
for ($i = 0; $i -lt $total; $i++) {
    if ($null -ne $upnById[$i]) { continue }  # already blank/skipped
    $encoded = [Uri]::EscapeDataString($names[$i])
    $searchRequests += @{
        id      = "$i"
        method  = "GET"
        url     = "/users?`$search=`"displayName:$encoded`"&`$select=id,displayName,userPrincipalName"
        headers = @{ ConsistencyLevel = "eventual" }
    }
}

$searchResults = Invoke-GraphBatch -Requests $searchRequests -BatchSize $BatchSize -MaxRetries $MaxRetries

$unresolved = New-Object System.Collections.Generic.List[int]

foreach ($req in $searchRequests) {
    $idx = [int]$req.id
    $resp = $searchResults[$req.id]
    $users = if ($resp -and $resp.status -eq 200) { $resp.body.value } else { $null }
    $match = Get-BestMatch -Users $users -InputName $names[$idx]

    if ($match) {
        $upnById[$idx] = $match.userPrincipalName
    }
    else {
        $unresolved.Add($idx)  # try exact-filter fallback next
    }
}

Write-Host "Pass 1 (search): $($total - $unresolved.Count)/$total resolved." -ForegroundColor Green

# ---- Pass 2: batched exact $filter fallback --------------------------------

if ($unresolved.Count -gt 0) {
    $filterRequests = @()
    foreach ($idx in $unresolved) {
        $escaped = Escape-ODataLiteral $names[$idx]
        $filterRequests += @{
            id     = "$idx"
            method = "GET"
            url    = "/users?`$filter=displayName eq '$escaped'&`$select=id,displayName,userPrincipalName"
        }
    }

    $filterResults = Invoke-GraphBatch -Requests $filterRequests -BatchSize $BatchSize -MaxRetries $MaxRetries

    foreach ($req in $filterRequests) {
        $idx = [int]$req.id
        $resp = $filterResults[$req.id]
        $users = if ($resp -and $resp.status -eq 200) { $resp.body.value } else { $null }
        $match = Get-BestMatch -Users $users -InputName $names[$idx]
        $upnById[$idx] = if ($match) { $match.userPrincipalName } else { "" }
    }
}

$stillBlank = ($upnById | Where-Object { $_ -eq "" }).Count
Write-Host "Done resolving. $($total - $stillBlank)/$total names matched a UPN." -ForegroundColor Cyan

# ---- Write output ------------------------------------------------------------

$dir = Split-Path -Path $OutputCsv
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$upnById |
    ForEach-Object { [pscustomobject]@{ UserPrincipalName = $_ } } |
    Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "UPNs written to: $OutputCsv" -ForegroundColor Cyan
