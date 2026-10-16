$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $apiKey = (Get-Content ./bugzilla-api-user.pl -Raw | docker compose exec -T bugzilla perl -)
    if ($LASTEXITCODE -ne 0 -or !$apiKey -or $apiKey.Trim() -notmatch '^[a-zA-Z0-9]+$') { throw 'No se pudo generar la clave de Bugzilla.' }
    $apiKey = $apiKey.Trim()
    $secretFile = Join-Path $env:APPDATA 'Microsoft\UserSecrets\29d62801-1fe2-4114-8818-ad1d23f5787d\secrets.json'
    $settings = Get-Content $secretFile -Raw | ConvertFrom-Json -AsHashtable
    $settings['Bugzilla:ApiKey'] = $apiKey
    $settings['Bugzilla:BaseUrl'] = 'http://localhost:8080'
    $settings | ConvertTo-Json -Depth 20 | Set-Content $secretFile
    $check = Invoke-RestMethod http://localhost:8080/rest/user/portafolio-api%40portafolio.local -Headers @{ 'X-BUGZILLA-API-KEY' = $apiKey } -TimeoutSec 15
    if ($check.users[0].name -ne 'portafolio-api@portafolio.local') { throw 'La identidad API no coincide.' }
    Write-Host 'Clave Bugzilla guardada; identidad API verificada.'
} finally { Pop-Location }

