# Script query stock ke OAA API
# Konversi dari curl ke PowerShell (Invoke-RestMethod)
#
# Cara pakai (1 nilai):
#   .\oaa_stock_query.ps1 -ArticleId "8000044321" -SourceId "SS20"
#
# Cara pakai (banyak nilai, pisahkan dengan koma):
#   .\oaa_stock_query.ps1 -ArticleId "8000044321","8000044322" -SourceId "SS20","SS21"
#
# Kalau tiap artikel unit-nya beda, isi -Unit sejumlah -ArticleId (urutan harus sama):
#   .\oaa_stock_query.ps1 -ArticleId "8000044321","8000044322" -SourceId "SS20" -Unit "EA","PC"

param(
    [Parameter(Mandatory = $true)]
    [string[]]$ArticleId,

    [Parameter(Mandatory = $true)]
    [string[]]$SourceId,

    [string[]]$Unit = @("EA")
)

$url = "https://ping.erajaya.com/msoaa/api/v1/oaa-stock-query"

$headers = @{
    "dbCode"        = "oaa-prod"
    "Authorization" = "Basic QXplYzpTM3J2aXMxbnQzcm40bA=="
    "Content-Type"  = "application/json"
    "Cookie"        = "saplb_*=(J2EE5060320)5060351"
}

# Susun ARTICLES: kalau -Unit diisi sejumlah -ArticleId, pasangkan 1-1;
# kalau cuma 1 nilai -Unit, dipakai sama untuk semua artikel.
$articles = @()
for ($i = 0; $i -lt $ArticleId.Count; $i++) {
    if ($Unit.Count -eq $ArticleId.Count) {
        $unitValue = $Unit[$i]
    } else {
        $unitValue = $Unit[0]
    }
    $articles += @{
        ARTICLE_ID = $ArticleId[$i]
        UNIT       = $unitValue
    }
}

$sources = @()
foreach ($s in $SourceId) {
    $sources += @{ SOURCE_ID = $s }
}

$body = @{
    DATA = @{
        IDENTIFICATION = @{
            INTERFACE_NAME  = "stockquery"
            PARTNERS        = "300"
            MSGID_EXTERNAL  = [guid]::NewGuid().ToString()
            ORDER_REF       = ""
            SYSTEM_ORIGIN   = ""
        }
        ARTICLES = $articles
        SOURCES  = $sources
    }
} | ConvertTo-Json -Depth 10

try {
    $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json"
    Write-Host "Berhasil:" -ForegroundColor Green
    $response | ConvertTo-Json -Depth 10
}
catch {
    Write-Host "Gagal:" -ForegroundColor Red
    Write-Host $_.Exception.Message

    if ($_.Exception.Response) {
        $stream = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $errorBody = $reader.ReadToEnd()
        Write-Host "Detail response error:"
        Write-Host $errorBody
    }
}
