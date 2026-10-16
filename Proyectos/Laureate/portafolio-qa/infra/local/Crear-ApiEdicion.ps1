$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $credentialFile = Join-Path $PSScriptRoot 'credentials.local.txt'
    if (!(Select-String -LiteralPath $credentialFile -SimpleMatch '=== API SQUASH TM - LECTURA Y EDICION ===' -Quiet)) {
    $headers = @{ Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('admin:admin')) }
    $request = @{ name = 'Portafolio QA lectura y edicion 2026-10-16'; permissions = 'READ_WRITE'; expiry_date = '2027-10-16' } | ConvertTo-Json
    $squash = Invoke-RestMethod -Method Post -Uri 'http://localhost:8090/squash/api/rest/latest/tokens' -Headers $headers -ContentType application/json -Body $request -TimeoutSec 30
    if (!$squash.generated_token) { throw 'Squash no devolvio el token.' }
    @('', '=== API SQUASH TM - LECTURA Y EDICION ===', 'Base URL: http://localhost:8090/squash/api/rest/latest/', 'Usuario: admin', "Token ID: $($squash.id)", "Permisos: $($squash.permissions)", "Vencimiento: $($squash.expiry_date)", "Authorization: Bearer $($squash.generated_token)") | Add-Content $credentialFile
    if ($squash.permissions -ne 'READ_WRITE') { throw 'El token no tiene READ_WRITE.' }
    $null = Invoke-RestMethod 'http://localhost:8090/squash/api/rest/latest/projects' -Headers @{ Authorization = "Bearer $($squash.generated_token)" } -TimeoutSec 30
    Write-Host 'Squash: token READ_WRITE creado, guardado y autenticado.'
    }
    $raw = Get-Content ./bugzilla-editor-api.pl -Raw | docker compose exec -T bugzilla perl -
    if ($LASTEXITCODE -ne 0) { throw 'Fallo la creacion de la API Bugzilla.' }
    $bugzilla = $raw | ConvertFrom-Json
    if (!$bugzilla.api_key) { throw 'Bugzilla no devolvio la clave.' }
    @('', '=== API BUGZILLA - LECTURA Y EDICION ===', 'Base URL: http://localhost:8080/rest/', "Usuario: $($bugzilla.login)", 'Permisos de cuenta: editbugs, canconfirm (lectura, creacion y edicion de defectos accesibles)', 'Vencimiento: sin vencimiento automatico; revocable desde Bugzilla', "X-BUGZILLA-API-KEY: $($bugzilla.api_key)", 'La clave hereda los permisos de la cuenta; no administra usuarios ni productos.') | Add-Content $credentialFile
    $identity = Invoke-RestMethod 'http://localhost:8080/rest/user/portafolio-editor-api%40portafolio.local' -Headers @{ 'X-BUGZILLA-API-KEY' = $bugzilla.api_key } -TimeoutSec 30
    if ($identity.users[0].name -ne $bugzilla.login) { throw 'Identidad Bugzilla inesperada.' }
    Add-Content $credentialFile "`nValidacion: autenticacion REST comprobada en ambas API; Squash READ_WRITE y grupos Bugzilla editbugs/canconfirm verificados."
    Write-Host 'Bugzilla: clave creada, guardada y autenticada; editbugs/canconfirm verificados.'
} finally { Pop-Location }

