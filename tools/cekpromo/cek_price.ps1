# ==============================================================================
# Script Name : Get-PromoPrice.ps1
# Description : Ambil data promo SRP price dengan parameter dinamis & logging
# ==============================================================================

[CmdletBinding()]
param (
    # Parameter Dinamis
    [Parameter(Mandatory = $false)]
    [string]$BUCode = "AZ02",

    [Parameter(Mandatory = $false)]
    [string]$ChannelId = "50",

    [Parameter(Mandatory = $false)]
    [string]$SKU = "8000003212",

    # Parameter Tetap / Default
    [Parameter(Mandatory = $false)]
    [string]$MemberGroup = "00",

    [Parameter(Mandatory = $false)]
    [string]$PromoService = "0",

    [Parameter(Mandatory = $false)]
    [int]$Qty = 1,

    # Header khusus
    [Parameter(Mandatory = $false)]
    [string]$HeaderToken = "Cookie: __cf_bm=HJ.QSYyzYZZ29Pube5EI3H6vgDzEA1ewdn5g6t0OLs0-1786529827.2531698-1.0.1.1-4IlJn4DSNbJDwEHA_EcZeqc5OxgNS9X5FaPAehkTMVfQHdDp5NiGeZWQZIGvDwCRpsntLLUBD6J4dQfJcZceO.M6wR_AKgUx4f4_GDsrkAsS7iXBBzlnu3EY43AZdGq5",

    # File Output Log
    [Parameter(Mandatory = $false)]
    [string]$LogFile = "$PSScriptRoot\promo_api_execution.log"
)

# Fungsi Logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    if ($Level -eq "ERROR") {
        Write-Host $logEntry -ForegroundColor Red
    } elseif ($Level -eq "SUCCESS") {
        Write-Host $logEntry -ForegroundColor Green
    } else {
        Write-Host $logEntry -ForegroundColor Cyan
    }
    
    Add-Content -Path $LogFile -Value $logEntry
}

# 1. Menyusun URL & Dynamic Query Parameters
$baseUrl = "https://erpi.eraspace.com/api/v1/promo-srp-price/item"
$queryParams = "bucode=$BUCode&channelid=$ChannelId&membergroup=$MemberGroup&promoservice=$PromoService&qty=$Qty&sku=$SKU"
$fullUrl = "$baseUrl`?$queryParams"

# 2. Menyusun Header
# Catatan: Header tanpa key/value pasif dalam cURL dimasukkan sebagai Custom Header
$headers = @{
    "Header-Token" = $HeaderToken  # Sesuaikan key header ini jika API membutuhkan nama spesifik (misal: Authorization)
}

Write-Log "Memulai Request ke URL: $fullUrl" "INFO"

try {
    # 3. Eksekusi API
    $response = Invoke-RestMethod -Uri $fullUrl -Method Get -Headers $headers -ErrorAction Stop

    Write-Log "BERHASIL! Data berhasil diambil untuk SKU: $SKU" "SUCCESS"
    
    # Tampilkan Respon JSON
    $response

} catch {
    $errorDetails = $_.Exception.Message
    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
        $errorDetails = "HTTP Status Code $statusCode - " + $_.Exception.Message
    }

    Write-Log "GAGAL! Detail Error: $errorDetails" "ERROR"
}