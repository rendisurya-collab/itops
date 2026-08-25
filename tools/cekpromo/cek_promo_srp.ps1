# ==============================================================================
# Script PowerShell untuk cek Promo SRP Price via API
#
# Cara pakai:
#   .\cek_promo_srp.ps1 -SKU "8100102377" -BUCode "AZ02"
#   .\cek_promo_srp.ps1 -SKU "8100102377" -BUCode "AZ02" -ChannelId "50" -MemberGroup "00"
#
# Output:
#   - Data promo SRP price ditampilkan di console (format JSON)
# ==============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$SKU,

    [Parameter(Mandatory = $true)]
    [string]$BUCode,

    [Parameter(Mandatory = $false)]
    [string]$ChannelId = "50",

    [Parameter(Mandatory = $false)]
    [string]$MemberGroup = "00"
)

$ApiUrl = "https://erpi.eraspace.com/api/v1/promo-srp-price/item"
$Cookie = "__cf_bm=NgLkZpEt34ziuG9W7LcEMGPCPSHf2djP9iIIgkeXO6E-1787595049.3498607-1.0.1.1-0JqP.JO.klFdLdHU_7hDoGy9rtjJouewDrH42t62YTIkkXTDwL4JN3Ul9vgU4nA47qndLI9Ko_Eqv3vF4aTI3Mildpi5WVcD1gUv.SQNHpdG4GTLTo0JuCsiUQ99To3S"

$fullUrl = "${ApiUrl}?bucode=${BUCode}&channelid=${ChannelId}&membergroup=${MemberGroup}&promoservice=0&qty=1&sku=${SKU}"

$headers = @{
    "Cookie" = $Cookie
}

try {
    $response = Invoke-RestMethod -Uri $fullUrl -Method Get -Headers $headers -ErrorAction Stop
    $response | ConvertTo-Json -Depth 10
}
catch {
    $statusCode = "N/A"
    $errorBody = $_.Exception.Message

    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $errorBody = $reader.ReadToEnd()
        } catch {}
    }

    Write-Host "Gagal (HTTP $statusCode):" -ForegroundColor Red
    Write-Host $errorBody
}
