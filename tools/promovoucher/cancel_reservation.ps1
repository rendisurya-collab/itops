# ============================================================
# Script PowerShell untuk CANCEL / hapus reservasi secara bulk via API
#
# Cara pakai:
#   1. Simpan file Excel sumber data sebagai CSV terlebih dahulu:
#        Excel -> File -> Save As -> CSV UTF-8 (Comma delimited) (*.csv)
#      Kolom yang dibutuhkan: businessUnitCode, transactionNumber, itemCode, orderDate
#   2. Buka PowerShell, masuk ke folder file ini & CSV disimpan, contoh:
#        cd C:\Users\NamaKamu\Downloads
#   3. Jalankan:
#        .\cancel_reservation.ps1 -CsvFile "data_delete_reservation.csv"
#
#   Jika muncul error "execution of scripts is disabled", jalankan dulu:
#        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#
# Output:
#   - Log lengkap tersimpan di cancel_log_<timestamp>.csv
#   - Ringkasan sukses/gagal ditampilkan di akhir
# ============================================================

param(
    [string]$CsvFile = "data_delete_reservation.csv"
)

# Paksa gunakan TLS 1.2 (mencegah error koneksi tidak jelas di PowerShell lama)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ====== KONFIGURASI ======
$ApiUrl = "https://erpi.eraspace.com/api/v1/pos/hub/main"

# Nilai-nilai FIXED (sesuai contoh JSON, sama untuk semua request)
$ClientId = "ERAFONEDOTCOM"
$ClientSignature = "6f19a056e10b367dc354cbec47de2fd93fad1c25f505517bf8473529e2959073"
$Quantity = 1
$Price = "0"
$OrderStatus = "cancel"
# ==========================

if (-not (Test-Path $CsvFile)) {
    Write-Host "Error: file '$CsvFile' tidak ditemukan." -ForegroundColor Red
    Write-Host "Pemakaian: .\cancel_reservation.ps1 -CsvFile 'data_delete_reservation.csv'"
    exit 1
}

$epochOrigin = Get-Date -Date "1970-01-01 00:00:00Z"

function ConvertTo-UnixEpoch($dateValue) {
    # Coba parse berbagai kemungkinan format tanggal dari Excel/CSV
    $parsed = $null
    if ($dateValue -is [datetime]) {
        $parsed = $dateValue
    }
    elseif ([string]::IsNullOrWhiteSpace($dateValue)) {
        return $null
    }
    elseif ($dateValue -match '^\d{9,10}$') {
        # sudah dalam bentuk unix epoch
        return [int64]$dateValue
    }
    else {
        [datetime]$tmp = Get-Date
        if ([datetime]::TryParse($dateValue, [ref]$tmp)) {
            $parsed = $tmp
        }
    }

    if ($null -eq $parsed) {
        return $null
    }
    return [int64]([datetime]$parsed - $epochOrigin).TotalSeconds
}

$rows = Import-Csv -Path $CsvFile

$total = 0
$successCount = 0
$failedCount = 0
$results = @()
$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "cancel_log_$logTimestamp.csv"

Write-Host "Memulai proses cancel reservasi dari file: $CsvFile"
Write-Host "----------------------------------------------"

foreach ($row in $rows) {
    $businessUnitCode = ($row.businessUnitCode).ToString().Trim()
    $transactionNumber = ($row.transactionNumber).ToString().Trim()
    $itemCode = ($row.itemCode).ToString().Trim()
    $orderDateEpoch = ConvertTo-UnixEpoch $row.orderDate

    if ([string]::IsNullOrWhiteSpace($transactionNumber)) {
        continue
    }

    $total++

    $nowUtc = (Get-Date).ToUniversalTime()
    $timestamps = [string][int64]($nowUtc - $epochOrigin).TotalSeconds
    $cancelDate = $nowUtc.ToString("yyyyMMddHHmmss")

    Write-Host "[$total] Cancel reservasi - transactionNumber: $transactionNumber, itemCode: $itemCode ..." -NoNewline

    $body = @{
        procedureCode = "O2O - CANCEL ORDER"
        parametersIn = @{
            clientId          = $ClientId
            timestamps        = $timestamps
            businessUnitCode  = $businessUnitCode
            clientSignature   = $ClientSignature
            transactionNumber = $transactionNumber
            orderDate         = [string]$orderDateEpoch
            cancelDate        = $cancelDate
            orderStatus       = $OrderStatus
        }
        childProcedure = @(
            @{
                procedureCode = "O2O - CANCEL ORDER ITEM"
                parametersIn = @{
                    transactionNumber = $transactionNumber
                    itemCode          = $itemCode
                    quantity          = $Quantity
                    price             = $Price
                }
            }
        )
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $ApiUrl -Method POST `
            -ContentType "application/json" -Body $body -ErrorAction Stop
        $statusCode = 200
        $responseBody = ($response | ConvertTo-Json -Depth 10 -Compress) -replace ",", ";"
        $status = "SUCCESS"
        $successCount++
        Write-Host "  -> Sukses (HTTP $statusCode)" -ForegroundColor Green
    }
    catch {
        $status = "FAILED"
        $failedCount++
        if ($_.Exception.Response) {
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { $statusCode = "N/A" }
        } else {
            $statusCode = "N/A"
        }

        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $responseBody = $_.ErrorDetails.Message -replace "`r`n|`n", " " -replace ",", ";"
        } else {
            $responseBody = ($_.Exception.Message + " | Detail: " + $_.Exception.GetType().FullName) -replace ",", ";"
        }
        Write-Host "  -> Gagal (HTTP $statusCode): $responseBody" -ForegroundColor Red
    }

    $results += [PSCustomObject]@{
        transactionNumber = $transactionNumber
        itemCode          = $itemCode
        businessUnitCode  = $businessUnitCode
        orderDate         = $orderDateEpoch
        cancelDate        = $cancelDate
        timestamps        = $timestamps
        url_hit           = $ApiUrl
        request_body      = ($body -replace "`r`n|`n", " " -replace ",", ";")
        http_status       = $statusCode
        response_body     = $responseBody
        status            = $status
    }

    Start-Sleep -Milliseconds 200
}

$results | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "===== SELESAI =====" -ForegroundColor Cyan
Write-Host "Total  : $total"
Write-Host "Sukses : $successCount" -ForegroundColor Green
Write-Host "Gagal  : $failedCount" -ForegroundColor Red
Write-Host "Log lengkap tersimpan di: $logFile"
