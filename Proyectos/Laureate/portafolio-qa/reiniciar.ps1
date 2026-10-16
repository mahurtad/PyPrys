# Detiene la app en el puerto 5011, recompila y la vuelve a lanzar en esta misma ventana.
# Uso: .\reiniciar.ps1
Set-Location $PSScriptRoot
Get-NetTCPConnection -LocalPort 5011 -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
$env:ASPNETCORE_ENVIRONMENT = 'Development'
dotnet run --no-launch-profile --urls http://localhost:5011
