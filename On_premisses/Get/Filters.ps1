Get-ADUser -Identity "Account.Alias" -Properties msexchmailboxguid 
| Select-Object Name,msexchmailboxguid