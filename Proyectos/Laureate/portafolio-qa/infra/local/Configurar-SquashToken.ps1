$ErrorActionPreference = 'Stop'
$headers = @{ Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('admin:admin')) }
$body = @{ name = 'Portafolio QA local lectura'; permissions = 'READ'; expiry_date = '2027-10-16' } | ConvertTo-Json
try {
    $response = Invoke-RestMethod -Method Post -Uri http://localhost:8090/squash/api/rest/latest/tokens -Headers $headers -ContentType application/json -Body $body -TimeoutSec 15
    $secretFile = Join-Path $env:APPDATA 'Microsoft\UserSecrets\29d62801-1fe2-4114-8818-ad1d23f5787d\secrets.json'
    $token = $response.generated_token
    if (!$token) { $token = $response.token_value }
    if (!$token) { Write-Output ('Campos de respuesta: ' + ($response.PSObject.Properties.Name -join ', ')); exit 2 }
    $secrets = Get-Content $secretFile -Raw | ConvertFrom-Json -AsHashtable
    $secrets['Squash:Token'] = $token
    $secrets['Squash:BaseUrl'] = 'http://localhost:8090/squash'
    $secrets | ConvertTo-Json -Depth 20 | Set-Content $secretFile
    $projects = Invoke-RestMethod -Uri http://localhost:8090/squash/api/rest/latest/projects -Headers @{ Authorization = "Bearer $token" } -TimeoutSec 15
    Write-Output 'Token Squash local guardado y GET projects validado.'
} catch {
    Write-Output $_.Exception.Message
    if ($_.ErrorDetails.Message) { Write-Output $_.ErrorDetails.Message }
    exit 1
}

