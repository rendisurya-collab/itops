# ==============================================================================
# Script PowerShell untuk cek AWB JNE via API Jeanne
#
# Cara pakai:
#   .\tracking_jne.ps1 -OrderNumber "8402663858" -AWB "0157352600237230"
# ==============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$OrderNumber,

    [Parameter(Mandatory = $true)]
    [string]$AWB
)

$ApiUrl = "https://jeanne.eraspace.com/shippings/v2/tracking/order/oms/jne"

$headers = @{
    "authorization" = "Basic c2hpcHBpbmdiYXNpYzo3NmNkNDJlZTQzZTUxNTIzZTAzNTVjZDE3NTMxY2ZjZjQxYjE2MWNmZDJjNTgwNDJkZjkxZTVmODU1MDQwYTQx"
    "x-source"      = "eraspace"
    "x-platform"    = "omsservice"
    "Content-Type"  = "application/json"
}

$body = @{
    order_number = $OrderNumber
    awb          = $AWB
} | ConvertTo-Json

try {
    $response = Invoke-RestMethod -Uri $ApiUrl -Method Post -Headers $headers -Body $body -ErrorAction Stop
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
