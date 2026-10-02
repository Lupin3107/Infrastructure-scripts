# Build the OPATH filter safely with a single-quoted here-string to avoid interpolation
$filter = @'
(
    (CustomAttribute13 -eq 'Example')
    -and (Office -ne 'Example Local')
    -and (Office -ne 'Example Local 2')
    -and (Name -notlike 'SystemMailbox{*')
    -and (Name -notlike 'CAS_{*')
    -and (RecipientTypeDetails -ne 'MailboxPlan')
    -and (RecipientTypeDetails -ne 'DiscoveryMailbox')
    -and (RecipientTypeDetails -ne 'PublicFolderMailbox')
    -and (RecipientTypeDetails -ne 'ArbitrationMailbox')
    -and (RecipientTypeDetails -ne 'AuditLogMailbox')
    -and (RecipientTypeDetails -ne 'AuxAuditLogMailbox')
    -and (RecipientTypeDetails -ne 'SupervisoryReviewPolicyMailbox')
    -and (
        (CustomAttribute10 -like '*DIRETOR* EXAMPLE*') -or
        (CustomAttribute10 -like '*DIRETORA* EXAMPLE*') -or
        (CustomAttribute10 -like '*DIRETOR(A)* EXAMPLE*') -or
        (CustomAttribute10 -like '*GERENTE* EXAMPLE*')
    )
)
'@

# Optional: add the disabled-account clause if your environment exposes ExchangeUserAccountControl
# Uncomment the next line if supported
# $filter = $filter -replace '\)\s*$', " -and (ExchangeUserAccountControl -ne 'AccountDisabled'))"

# Show the final filter (helpful for debugging)
Write-Host "Using RecipientPreviewFilter: $filter"

try {
    if (Get-Command -Name Get-Recipient -ErrorAction SilentlyContinue) {
        Get-Recipient -RecipientPreviewFilter $filter -ResultSize Unlimited | Select-Object Name, Title
    } else {
        Write-Warning "Get-Recipient cmdlet not available in this session. Filter prepared but not executed."
    }
} catch {
    Write-Error "Error running Get-Recipient: $_"
}