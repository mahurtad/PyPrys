# Crea (una sola vez) la cuenta de servicio de Bugzilla con la que Portafolio crea productos y componentes por la API REST.
# Solo configura la cuenta: la funcionalidad de Portafolio no usa Docker, Perl ni SQL contra Bugzilla.
# La clave va a user-secrets (Bugzilla:AdminApiKey) y a credentials.local.txt (ignorado por Git); no se imprime.
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $secretFile = Join-Path $env:APPDATA 'Microsoft\UserSecrets\29d62801-1fe2-4114-8818-ad1d23f5787d\secrets.json'
    $settings = Get-Content $secretFile -Raw | ConvertFrom-Json -AsHashtable
    if ($settings['Bugzilla:AdminApiKey']) { Write-Host 'Bugzilla:AdminApiKey ya existe; no se crea otra clave.'; return }
    $raw = Get-Content ./bugzilla-productos-api.pl -Raw | docker compose exec -T bugzilla perl -
    if ($LASTEXITCODE -ne 0) { throw 'Fallo la creacion de la cuenta de servicio.' }
    $cuenta = $raw | ConvertFrom-Json
    if (!$cuenta.api_key) { throw 'Bugzilla no devolvio la clave.' }
    $settings['Bugzilla:AdminApiKey'] = $cuenta.api_key
    $settings | ConvertTo-Json -Depth 20 | Set-Content $secretFile
    @('', '=== API BUGZILLA - PRODUCTOS (cuenta de servicio de Portafolio) ===', 'Base URL: http://localhost:8080/rest/', "Usuario: $($cuenta.login)", 'Permisos de cuenta: editcomponents (crea productos y componentes)', "X-BUGZILLA-API-KEY: $($cuenta.api_key)", 'Guardada tambien en user-secrets como Bugzilla:AdminApiKey. En servidor: variable de entorno Bugzilla__AdminApiKey.') | Add-Content (Join-Path $PSScriptRoot 'credentials.local.txt')
    $user = Invoke-RestMethod "http://localhost:8080/rest/user/$([uri]::EscapeDataString($cuenta.login))" -Headers @{ 'X-BUGZILLA-API-KEY' = $cuenta.api_key } -TimeoutSec 30
    if ($user.users[0].name -ne $cuenta.login) { throw 'Identidad Bugzilla inesperada.' }
    Write-Host 'Cuenta de servicio de productos creada y autenticada; clave guardada en user-secrets. Reinicia Portafolio.'
} finally { Pop-Location }
