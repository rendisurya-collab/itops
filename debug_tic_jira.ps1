# Script test debug: query project TIC langsung ke Jira API
# Ganti email dan token sesuai isi file .env kamu (JIRA_EMAIL dan JIRA_API_TOKEN)

$user = "isi_email_jira_kamu@company.com"
$token = "isi_JIRA_API_TOKEN_dari_env"
$baseUrl = "isi_JIRA_BASE_URL_dari_env"   # contoh: https://namaperusahaan.atlassian.net

$pair = "$user`:$token"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($pair)
$base64 = [Convert]::ToBase64String($bytes)
$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json"
}

Write-Host "1) Cek siapa akun ini di Jira:" -ForegroundColor Cyan
try {
    $me = Invoke-RestMethod -Uri "$baseUrl/rest/api/3/myself" -Headers $headers
    Write-Host "   OK - login sebagai: $($me.displayName) ($($me.emailAddress))" -ForegroundColor Green
} catch {
    Write-Host "   GAGAL login:" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)"
}

Write-Host ""
Write-Host "2) Cek project TIC ada dan bisa diakses:" -ForegroundColor Cyan
try {
    $proj = Invoke-RestMethod -Uri "$baseUrl/rest/api/3/project/TIC" -Headers $headers
    Write-Host "   OK - project ditemukan: $($proj.name) (key: $($proj.key), id: $($proj.id))" -ForegroundColor Green
} catch {
    Write-Host "   GAGAL akses project TIC:" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)"
    if ($_.Exception.Response) {
        $stream = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        Write-Host "   Detail:" $reader.ReadToEnd()
    }
}

Write-Host ""
Write-Host "3) Cek search JQL langsung ke project TIC:" -ForegroundColor Cyan
$body = @{
    jql        = 'project = "TIC" ORDER BY updated DESC'
    fields     = @("summary", "status")
    maxResults = 5
} | ConvertTo-Json

try {
    $result = Invoke-RestMethod -Uri "$baseUrl/rest/api/3/search/jql" -Method Post -Headers $headers -Body $body -ContentType "application/json"
    Write-Host "   Jumlah issue ditemukan: $($result.issues.Count)" -ForegroundColor Green
    $result.issues | ForEach-Object { Write-Host "   - $($_.key): $($_.fields.summary)" }
} catch {
    Write-Host "   GAGAL search:" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)"
    if ($_.Exception.Response) {
        $stream = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        Write-Host "   Detail:" $reader.ReadToEnd()
    }
}
