$skuMap = @{
    "00a000a0-0000-a000-0000-00a0a000a000" = "Microsoft 365 E3"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Office 365 E1"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Office 365 E1 Plus"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Microsoft 365 F3"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Office 365 F3"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Microsoft F1"
    "00a000a0-0000-a000-0000-00a0a000a000" = "Exchange Archiving"
}

$upn = Read-Host "Digite o nome da conta"

if ([string]::IsNullOrWhiteSpace($upn)) {
    Write-Host "Nenhum UPN informado. Encerrando o script." -ForegroundColor Yellow
    exit
}

try {
    $licencasDetalhadas = Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop

    $licencas = foreach ($licenca in $licencasDetalhadas) {
        $skuId = $licenca.SkuId.ToString()

        if ($skuMap.ContainsKey($skuId)) {
            $skuMap[$skuId]
        }
        else {
            $skuId
        }
    }

    if ($licencas.Count -eq 0) {
        Write-Host "! Nenhuma licença encontrada para $upn" -ForegroundColor Yellow
    }
    else {
        Write-Host "Licenças encontradas: $($licencas -join ', ')"
    }
}
catch {
    Write-Host "✕ Erro ao consultar o usuário: $upn" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor DarkRed
}