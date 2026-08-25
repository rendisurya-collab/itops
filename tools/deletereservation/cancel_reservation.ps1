# ============================================================
# Script PowerShell untuk CANCEL / hapus reservasi secara bulk via API
#
# Cara pakai:
#   1. Simpan file Excel sumber data sebagai CSV terlebih dahulu:
#        Excel -> File -> Save As -> CSV UTF-8 (Comma delimited) (*.csv)
#      Urutan kolom yang diharapkan (header row):
#        businessUnitCode, transactionNumber, itemCode, itemCode, ..., qty, orderDate
#      Catatan: kolom itemCode boleh muncul lebih dari 1x (berdampingan) untuk
#      transaksi yang punya banyak item. Kolom qty diisi jumlah quantity yang
#      ingin di-cancel (angka bulat, default 1 jika kosong).
#      orderDate WAJIB jadi kolom PALING AKHIR.
#   2. Buka PowerShell, masuk ke folder file ini & CSV disimpan, contoh:
#        cd C:\Users\NamaKamu\Downloads
#   3. Jalankan:
#        .\cancel_reservation.ps1 -CsvFile "delete_reservation.csv"
#      Atau dengan qty default (berlaku untuk semua baris yang qty-nya kosong):
#        .\cancel_reservation.ps1 -CsvFile "delete_reservation.csv" -DefaultQty 2
#
#   Jika muncul error "execution of scripts is disabled", jalankan dulu:
#        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#
# Output:
#   - Log lengkap tersimpan di cancel_log_<timestamp>.csv
#   - Ringkasan sukses/gagal ditampilkan di akhir
# ============================================================

param(
    [string]$CsvFile = "delete_reservation.csv",
    [int]$DefaultQty = 1
)

# Paksa gunakan TLS 1.2 (mencegah error koneksi tidak jelas di PowerShell lama)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ====== KONFIGURASI ======
$ApiUrl = "https://erpi.eraspace.com/api/v1/pos/hub/main"

# Nilai-nilai FIXED (sesuai contoh JSON, sama untuk semua request)
$ClientId = "ERAFONEDOTCOM"
$ClientSignature = "6f19a056e10b367dc354cbec47de2fd93fad1c25f505517bf8473529e2959073"
$Price = "0"
$OrderStatus = "cancel"
# ==========================

if (-not (Test-Path $CsvFile)) {
    Write-Host "Error: file '$CsvFile' tidak ditemukan." -ForegroundColor Red
    Write-Host "Pemakaian: .\cancel_reservation.ps1 -CsvFile 'delete_reservation.csv'"
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

# Baca CSV berbasis POSISI kolom (bukan nama), karena file punya beberapa
# kolom "itemCode" dengan nama sama berdampingan. Urutan kolom yang diharapkan:
#   [0] businessUnitCode, [1] transactionNumber, [2..n-3] itemCode(s), [n-2] qty, [n-1] orderDate
# Jika qty tidak ada atau kosong, gunakan $DefaultQty.
$allLines = Get-Content -Path $CsvFile -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

if ($allLines.Count -lt 2) {
    Write-Host "Error: file CSV tidak berisi data (hanya header atau kosong)." -ForegroundColor Red
    exit 1
}

# Deteksi apakah header punya kolom 'qty'
$headerFields = ($allLines[0] -split ',') | ForEach-Object { $_.Trim().Trim('"').ToLower() }
$hasQtyColumn = $headerFields -contains "qty"

$parsedRows = @()
foreach ($line in ($allLines | Select-Object -Skip 1)) {
    $fields = $line -split ',' | ForEach-Object { $_.Trim().Trim('"') }
    if ($fields.Count -lt 4) { continue }

    $lastIndex = $fields.Count - 1

    if ($hasQtyColumn) {
        # Format: businessUnitCode, transactionNumber, itemCode(s)..., qty, orderDate
        $qtyIndex = $lastIndex - 1
        $qtyRaw = $fields[$qtyIndex]
        $qty = if ($qtyRaw -match '^\d+$') { [int]$qtyRaw } else { $DefaultQty }
        $itemCodes = @($fields[2..($qtyIndex - 1)] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } else {
        # Format lama tanpa kolom qty: businessUnitCode, transactionNumber, itemCode(s)..., orderDate
        $qty = $DefaultQty
        $itemCodes = @($fields[2..($lastIndex - 1)] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $parsedRows += [PSCustomObject]@{
        businessUnitCode  = $fields[0]
        transactionNumber = $fields[1]
        itemCodes         = $itemCodes
        quantity          = $qty
        orderDate         = $fields[$lastIndex]
    }
}

$total = 0
$successCount = 0
$failedCount = 0
$results = @()


$logTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFolder = "D:\mybot\tools\deletereservation\logs"

# Buat folder-nya otomatis kalau belum ada
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}
$logFile = Join-Path $LogFolder "cancel_log_$logTimestamp.csv"

Write-Host "Memulai proses cancel reservasi dari file: $CsvFile"
Write-Host "----------------------------------------------"

foreach ($row in $parsedRows) {
    $transactionNumber = $row.transactionNumber

    if ([string]::IsNullOrWhiteSpace($transactionNumber)) {
        continue
    }

    $businessUnitCode = $row.businessUnitCode
    $orderDateEpoch = ConvertTo-UnixEpoch $row.orderDate
    $qty = $row.quantity

    # Kumpulkan semua itemCode di baris ini jadi array childProcedure
    $itemCodesList = @()
    $childProcedures = @()
    foreach ($itemCode in $row.itemCodes) {
        $itemCodesList += $itemCode
        $childProcedures += @{
            procedureCode = "O2O - CANCEL ORDER ITEM"
            parametersIn = @{
                transactionNumber = $transactionNumber
                itemCode          = $itemCode
                quantity          = $qty
                price             = $Price
            }
        }
    }

    if ($childProcedures.Count -eq 0) {
        continue
    }

    $total++

    $nowUtc = (Get-Date).ToUniversalTime()
    $timestamps = [string][int64]($nowUtc - $epochOrigin).TotalSeconds
    $cancelDate = $nowUtc.ToString("yyyyMMddHHmmss")

    $itemCodesDisplay = $itemCodesList -join ", "
    Write-Host "[$total] Cancel reservasi - transactionNumber: $transactionNumber, itemCode(s): $itemCodesDisplay, qty: $qty ..." -NoNewline

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
        childProcedure = $childProcedures
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $ApiUrl -Method POST `
            -ContentType "application/json" -Body $body -ErrorAction Stop
        $statusCode = 200
        $rawResponseJson = ($response | ConvertTo-Json -Depth 10 -Compress)
        $responseBody = $rawResponseJson -replace ",", ";"

        # Cek apakah ada MSG_DESC di response, meski HTTP 200 bisa saja
        # sebenarnya gagal secara logika (misal "External ID xxx not exists")
        $msgDesc = $null
        if ($rawResponseJson -match '"MSG_DESC"\s*:\s*"([^"]*)"') {
            $msgDesc = $matches[1]
        }

        if ($msgDesc -and ($msgDesc -match '(?i)not exists|not found|error|fail|invalid|gagal')) {
            $status = "FAILED (LOGICAL)"
            $failedCount++
            Write-Host "  -> Gagal secara logika (HTTP $statusCode): $msgDesc" -ForegroundColor Yellow
        } else {
            $status = "SUCCESS"
            $successCount++
            Write-Host "  -> Sukses (HTTP $statusCode)" -ForegroundColor Green
        }
    }
    catch {
        $status = "FAILED"
        $failedCount++
        $msgDesc = $null
        if ($_.Exception.Response) {
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { $statusCode = "N/A" }
        } else {
            $statusCode = "N/A"
        }

        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $responseBody = $_.ErrorDetails.Message -replace "`r`n|`n", " " -replace ",", ";"
            if ($_.ErrorDetails.Message -match '"MSG_DESC"\s*:\s*"([^"]*)"') {
                $msgDesc = $matches[1]
            }
        } else {
            $responseBody = ($_.Exception.Message + " | Detail: " + $_.Exception.GetType().FullName) -replace ",", ";"
        }
        Write-Host "  -> Gagal (HTTP $statusCode): $responseBody" -ForegroundColor Red
    }

    $results += [PSCustomObject]@{
        transactionNumber = $transactionNumber
        itemCodes         = $itemCodesDisplay
        quantity          = $qty
        businessUnitCode  = $businessUnitCode
        orderDate         = $orderDateEpoch
        cancelDate        = $cancelDate
        timestamps        = $timestamps
        url_hit           = $ApiUrl
        request_body      = ($body -replace "`r`n|`n", " " -replace ",", ";")
        http_status       = $statusCode
        msg_desc          = $msgDesc
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
