$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    docker compose up -d
    if ($LASTEXITCODE -ne 0) { throw 'No se pudieron iniciar los servicios. Inicia Docker Desktop y vuelve a ejecutar.' }
    docker compose ps
    Write-Host 'Squash TM: http://localhost:8090/squash'
    Write-Host 'Bugzilla: http://localhost:8080'
} finally { Pop-Location }
