# ============================================================
# Script PowerShell untuk release/hapus promo customer secara bulk via API
#
# Cara pakai:
#   1. Buka PowerShell
#   2. Masuk ke folder tempat file ini & promo_id.csv disimpan, contoh:
#        cd C:\Users\NamaKamu\Downloads
#   3. Jalankan:
#        .\release_promo.ps1 -CsvFile "promo_id.csv"
#
#   Jika muncul error "execution of scripts is disabled", jalankan dulu
#   (sekali saja, di sesi PowerShell yang sama):
#        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   lalu ulangi perintah di langkah 3.
#
# Output:
#   - Log lengkap tersimpan di release_log_<timestamp>.csv
#   - Ringkasan sukses/gagal ditampilkan di akhir
# ============================================================

param(
    [string]$CsvFile = "promo_id.csv"
)

# ====== KONFIGURASI ======
$ApiBase = "https://sculptor.eraspace.com/promos/v1/promo/release-by-system"
$AuthHeaderName = "X-Auth-Signature"
$AuthHeaderValue = "e9738b21b981a6f33d096f51830fac27"
# ==========================

if (-not (Test-Path $CsvFile)) {
    Write-Host "Error: file '$CsvFile' tidak ditemukan." -ForegroundColor Red
    Write-Host "Pemakaian: .\release_promo.ps1 -CsvFile 'promo_id.csv'"
    exit 1
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "release_log_$timestamp.csv"

# Baca CSV (otomatis mendeteksi kolom header, misal "id")
$rows = Import-Csv -Path $CsvFile
$firstColumnName = ($rows | Get-Member -MemberType NoteProperty | Select-Object -First 1).Name

$total = 0
$successCount = 0
$failedCount = 0
$results = @()

Write-Host "Memulai proses release promo dari file: $CsvFile"
Write-Host "----------------------------------------------"

foreach ($row in $rows) {
    $id = ($row.$firstColumnName).ToString().Trim()

    if ([string]::IsNullOrWhiteSpace($id)) {
        continue
    }

    $total++
    Write-Host "[$total] Releasing promo ID: $id ..." -NoNewline

    $url = "$ApiBase/$id"
    $headers = @{ $AuthHeaderName = $AuthHeaderValue }

    try {
        $response = Invoke-WebRequest -Uri $url -Method POST -Headers $headers -UseBasicParsing -ErrorAction Stop
        $statusCode = $response.StatusCode
        $body = $response.Content -replace "`r`n|`n", " " -replace ",", ";"
        $status = "SUCCESS"
        $successCount++
        Write-Host "  -> Sukses (HTTP $statusCode)" -ForegroundColor Green
    }
    catch {
        $status = "FAILED"
        $failedCount++
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $body = $reader.ReadToEnd() -replace "`r`n|`n", " " -replace ",", ";"
            } catch {
                $body = $_.Exception.Message -replace ",", ";"
            }
        } else {
            $statusCode = "N/A"
            $body = $_.Exception.Message -replace ",", ";"
        }
        Write-Host "  -> Gagal (HTTP $statusCode): $body" -ForegroundColor Red
    }

    $results += [PSCustomObject]@{
        id            = $id
        http_status   = $statusCode
        response_body = $body
        status        = $status
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