<#
 Genera el paquete instalable de Portafolio QA (carpeta + .zip) en .\dist
 Uso (en la máquina de desarrollo):  .\Instalador\Construir-Instalador.ps1
 Requiere el SDK de .NET 8 y la herramienta dotnet-ef.
#>
$ErrorActionPreference = 'Stop'
$raiz = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $raiz 'dist'
$paquete = Join-Path $dist 'PortafolioQA-Instalador'
$work = Join-Path $env:TEMP ('pqa_build_' + [guid]::NewGuid().ToString('N'))

Write-Host '1/4 Copiando fuentes a una carpeta temporal (evita bloqueos de Google Drive / app en ejecución)...'
New-Item -ItemType Directory $work | Out-Null
# Se excluyen solo las carpetas de la RAÍZ (ruta completa). Con nombres sueltos, "dist" también excluiría wwwroot\libootstrap\dist,
# wwwroot\lib\jquery\dist, etc., y el sitio publicado quedaría sin estilos ni scripts.
robocopy $raiz $work /E /XD (Join-Path $raiz 'bin') (Join-Path $raiz 'obj') (Join-Path $raiz 'dist') (Join-Path $raiz '.git') /NFL /NDL /NJH /NJS /NP | Out-Null

Write-Host '2/4 Publicando (Release)...'
Remove-Item $paquete -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory (Join-Path $paquete 'app'), (Join-Path $paquete 'db') -Force | Out-Null
Push-Location $work
try {
    dotnet publish -c Release -o (Join-Path $paquete 'app') | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'dotnet publish falló.' }

    Write-Host '3/4 Generando el script SQL idempotente de migraciones...'
    dotnet ef migrations script --idempotent --configuration Release -o (Join-Path $paquete 'db\portafolio-qa.sql') | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'dotnet ef migrations script falló.' }
} finally { Pop-Location }

Write-Host '4/4 Armando el paquete...'
Copy-Item (Join-Path $PSScriptRoot 'Instalar.ps1'), (Join-Path $PSScriptRoot 'Utilidades.ps1'), (Join-Path $PSScriptRoot 'Instalar.cmd'), (Join-Path $PSScriptRoot 'LEEME.txt') $paquete
Get-ChildItem (Join-Path $paquete 'app') -Filter 'appsettings.Development.json' | Remove-Item -Force
# El paquete no debe llevar datos de un entorno concreto: se quitan la URL de Squash y la cadena de conexión de desarrollo.
# Sin cadena de conexión la aplicación arranca "pendiente" y se configura desde su asistente (/Setup).
$cfg = Join-Path (Join-Path $paquete 'app') 'appsettings.json'
$json = Get-Content $cfg -Raw -Encoding UTF8 | ConvertFrom-Json
if ($json.PSObject.Properties['Squash']) { $json.Squash.PSObject.Properties.Remove('BaseUrl') }
$json.PSObject.Properties.Remove('ConnectionStrings')
[IO.File]::WriteAllText($cfg, ($json | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($false)))
$zip = Join-Path $dist 'PortafolioQA-Instalador.zip'
Remove-Item $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $paquete '*') -DestinationPath $zip
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Paquete listo: $zip"
$iscc = @(foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, "$env:LOCALAPPDATA\Programs")) { Get-ChildItem $base -Filter 'Inno Setup*' -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | ForEach-Object { Join-Path $_.FullName 'ISCC.exe' } }) | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($iscc) {
    Remove-Item (Join-Path $dist 'PortafolioQA-Setup.exe') -Force -ErrorAction SilentlyContinue
    & $iscc (Join-Path $PSScriptRoot 'PortafolioQA.iss') | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'La compilación de Inno Setup falló (no se generó Setup.exe).' }
    Write-Host "Setup.exe listo: $(Join-Path $dist 'PortafolioQA-Setup.exe')"
} else {
    Write-Host 'Inno Setup no está instalado: se generó solo el .zip (Instalar.cmd). Instálelo y vuelva a ejecutar para obtener el Setup.exe.' -ForegroundColor Yellow
}
