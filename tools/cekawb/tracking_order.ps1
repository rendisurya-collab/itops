# ==============================================================================
# Script PowerShell untuk cek AWB / Tracking Order via API OMS
#
# Cara pakai:
#   .\tracking_order.ps1 -OrderNumber "3301352973" -Source "IBOX"
#
# Output:
#   - Detail tracking order ditampilkan di console (format JSON)
# ==============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$OrderNumber,

    [Parameter(Mandatory = $true)]
    [string]$Source
)

$ApiUrl = "https://erpi.eraspace.com/api/v1/oms/tracking/order"
$Cookie = "__cf_bm=9pJvt_cv9ezXcN2vikv0mDUfpZqKZgZytb9_UXECBts-1787596242.58195-1.0.1.1-sZNmD5Ut4qoO8qBRaiMF8ZG0jHx1gn.zsfZyPzT0mh1xCD4vcqeeczduxHrHvfs6q2VW4Ct4IYx8Ew6vNu5TjbdlGf1uCjIjaV8odm86fuSiwn.8hOhBsskMJUzS.uLf"

$body = @{
    orderNumber = $OrderNumber
    source      = $Source
} | ConvertTo-Json

$headers = @{
    "Content-Type" = "application/json"
    "Cookie"       = $Cookie
}

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
