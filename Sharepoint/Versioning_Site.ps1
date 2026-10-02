Connect-SPOService -Url "https://domain-admin.sharepoint.com"
$SiteURL = "https://domain.sharepoint.com/sites/Path"

Get-SPOListVersionPolicy `
    -Site $SiteURL `
    -List "Documentos" | FL VersioningEnabled, MajorVersionLimit

Get-SPOSiteFileVersionBatchDeleteJobProgress `
  -Identity $SiteURL | FL Status, RequestTimeInUTC, LastProcessTimeInUTC, MajorVersionLimit, FilesProcessed, VersionsProcessed,VersionsDeleted