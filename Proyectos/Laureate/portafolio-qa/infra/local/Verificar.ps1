$ErrorActionPreference = 'Stop'
$secretFile = Join-Path $env:APPDATA 'Microsoft\UserSecrets\29d62801-1fe2-4114-8818-ad1d23f5787d\secrets.json'
$config = Get-Content $secretFile -Raw | ConvertFrom-Json
if ($config.'Squash:BaseUrl' -ne 'http://localhost:8090/squash' -or $config.'Bugzilla:BaseUrl' -ne 'http://localhost:8080') { throw 'Las URL efectivas de user-secrets no son las instalaciones locales.' }
$squash = Invoke-RestMethod 'http://localhost:8090/squash/api/rest/latest/projects' -Headers @{ Authorization = "Bearer $($config.'Squash:Token')" } -TimeoutSec 30
Write-Host "OK Squash TM: API autenticada; proyectos accesibles: $($squash.page.totalElements)"
$apiKey = [Uri]::EscapeDataString($config.'Bugzilla:ApiKey')
try {
    $identity = Invoke-RestMethod "http://localhost:8080/rest/user/portafolio-api%40portafolio.local?api_key=$apiKey" -TimeoutSec 30
} catch {
    throw 'Fallo la consulta autenticada de Bugzilla. Comprueba el servicio y la clave de lectura.'
}
if ($identity.users[0].name -ne 'portafolio-api@portafolio.local') { throw 'Identidad Bugzilla inesperada.' }
$version = Invoke-RestMethod 'http://localhost:8080/rest/version' -TimeoutSec 30
Write-Host "OK Bugzilla: versión $($version.version), API autenticada."
$health = Invoke-RestMethod 'http://localhost:5011/healthz' -TimeoutSec 15
if ($health -ne 'ok') { throw 'Portafolio no está activo.' }
Write-Host 'OK Portafolio QA: aplicación y SQL Server disponibles.'
