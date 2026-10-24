# Inicia Squash y Bugzilla en Docker, y Portafolio en esta ventana.
# Requiere Docker Desktop iniciado. Uso: .\Iniciar-Todo.ps1
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'infra\local\Iniciar.ps1')

Write-Host 'Iniciando Portafolio: http://localhost:5011'
& (Join-Path $PSScriptRoot 'reiniciar.ps1')
