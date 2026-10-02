$licenses_assign = Get-MgUserLicenseDetail -UserId "example@domain.com"
$licenses_assign | Select-Object SkuId, SkuPartNumber, ServicePlans