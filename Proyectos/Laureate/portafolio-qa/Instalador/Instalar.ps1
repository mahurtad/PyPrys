<#
 Instalador de Portafolio QA en IIS. Se ejecuta desde Instalar.cmd o desde el Setup.exe (Inno Setup).
 Compatible con Windows PowerShell 5.1. No instala prerrequisitos: los valida y avisa.

 Con -Parametros (archivo clave=valor que escribe el asistente gráfico de Setup.exe) corre sin preguntar y borra el
 archivo al leerlo. Sin -Parametros pregunta por consola, como siempre. -Log guarda la salida en un archivo.
 Claves: Modo (nuevo|actualizar), Sitio, Carpeta, Host, Puerto, CarpetaClaves, CertThumb (huella o AUTOFIRMADO),
 AdminCorreo, AdminClave, SquashUrl, SquashToken.
 NO toca SQL Server: no crea bases, logins ni usuarios ni aplica migraciones. La conexión a la base se configura después,
 desde el asistente de la aplicación (/Setup); hasta entonces la aplicación queda instalada pero sin activar.
 Código de salida: 0 = ok (activa o pendiente de configurar), 1 = cancelado/falló, 2 = instalado pero el sitio no respondió.
#>
param(
    [string]$RutaApp = (Join-Path $PSScriptRoot 'app'),
    [string]$Parametros,
    [string]$Log
)
$ErrorActionPreference = 'Stop'

# La administración de IIS (WebAdministration, IIS:\) solo funciona en PowerShell de 64 bits. Si este proceso es de
# 32 bits en un Windows de 64 bits (p. ej. lo lanzó un instalador de 32 bits), se relanza con el PowerShell nativo.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $ps64 = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $ps64) {
        $args64 = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        foreach ($k in $PSBoundParameters.Keys) { $args64 += "-$k"; $args64 += [string]$PSBoundParameters[$k] }
        & $ps64 @args64
        exit $LASTEXITCODE
    }
}
$codigoSalida = 0

$P = @{}
$Silencioso = -not [string]::IsNullOrWhiteSpace($Parametros)
if ($Silencioso) {
    foreach ($linea in Get-Content $Parametros -Encoding UTF8) {
        $i = $linea.IndexOf('=')
        if ($i -gt 0) { $P[$linea.Substring(0, $i).Trim()] = $linea.Substring($i + 1) }
    }
    Remove-Item $Parametros -Force -ErrorAction SilentlyContinue
}
if ($Log) { Start-Transcript -Path $Log -Force | Out-Null }
function Terminar($codigo) { if ($Log) { try { Stop-Transcript | Out-Null } catch {} }; exit $codigo }

function Titulo($t) { Write-Host ''; Write-Host "== $t ==" -ForegroundColor Cyan }
function Pedir($msg, $def = '', $clave = $null) {
    if ($Silencioso) { if ($clave -and $P.ContainsKey($clave) -and -not [string]::IsNullOrWhiteSpace($P[$clave])) { return $P[$clave].Trim() } else { return $def } }
    $r = Read-Host ($msg + $(if ($def) { " [$def]" } else { '' })); if ([string]::IsNullOrWhiteSpace($r)) { $def } else { $r.Trim() }
}
function PedirSecreto($msg, $clave = $null) {
    if ($Silencioso) { if ($clave -and $P.ContainsKey($clave)) { return $P[$clave] } else { return '' } }
    $s = Read-Host $msg -AsSecureString
    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))
}
function Abortar($msg) { Write-Host ''; Write-Host "X $msg" -ForegroundColor Red; Write-Host 'Instalación cancelada.'; Terminar 1 }

# ---------- 1. Prerrequisitos ----------
Titulo 'Verificando requisitos del servidor'
$faltan = @()
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { $faltan += 'Ejecutar como Administrador.' }
if ($PSVersionTable.PSVersion.Major -lt 5) { $faltan += 'Windows PowerShell 5.1 o superior.' }
if (-not (Test-Path (Join-Path $RutaApp 'PortafolioQA.dll'))) { $faltan += "Falta la carpeta 'app' del paquete (PortafolioQA.dll)." }
if (-not (Get-Service W3SVC -ErrorAction SilentlyContinue)) { $faltan += 'IIS (rol Web Server) no está instalado. Instálelo con: Install-WindowsFeature Web-Server,Web-Mgmt-Tools (Server) o desde "Activar o desactivar características de Windows".' }
if (-not (Get-Module -ListAvailable WebAdministration)) { $faltan += 'Falta el módulo WebAdministration (Herramientas de administración de IIS / Web-Scripting-Tools).' }
$modulo = Join-Path $env:ProgramFiles 'IIS\Asp.Net Core Module\V2\aspnetcorev2.dll'
$runtime = Get-ChildItem (Join-Path $env:ProgramFiles 'dotnet\shared\Microsoft.AspNetCore.App') -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '8.*' }
if (-not (Test-Path $modulo) -or -not $runtime) { $faltan += 'Falta el .NET 8 Hosting Bundle (ASP.NET Core Module V2 + runtime 8). Descárguelo de https://dotnet.microsoft.com/download/dotnet/8.0, instálelo y ejecute "iisreset".' }
if ($faltan.Count -gt 0) {
    Write-Host 'No se puede continuar. Falta:' -ForegroundColor Red
    $faltan | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Abortar 'Resuelva los puntos anteriores y vuelva a ejecutar el instalador.'
}
Write-Host 'OK: IIS, WebAdministration y .NET 8 Hosting Bundle presentes.' -ForegroundColor Green
try {
    Import-Module WebAdministration -ErrorAction Stop
    if (-not (Get-PSDrive IIS -ErrorAction SilentlyContinue)) { throw "El proveedor IIS:\ no se cargó." }
    $null = Get-ChildItem 'IIS:\Sites' -ErrorAction Stop
    Write-Host 'OK: el proveedor IIS:\ responde.' -ForegroundColor Green
} catch {
    $bits = if ([Environment]::Is64BitProcess) { '64 bits' } else { '32 bits' }
    Write-Host ''
    Write-Host "X No se pudo usar IIS desde PowerShell (proceso de $bits): $($_.Exception.Message)" -ForegroundColor Red
    Write-Host '  Causas habituales:' -ForegroundColor Yellow
    Write-Host '   - Falta la característica "Scripts y herramientas de administración de IIS" (Web-Scripting-Tools) o la' -ForegroundColor Yellow
    Write-Host '     "Consola de administración de IIS" (Web-Mgmt-Console). En Windows Server:' -ForegroundColor Yellow
    Write-Host '       Install-WindowsFeature Web-Scripting-Tools,Web-Mgmt-Console' -ForegroundColor Yellow
    Write-Host '     En Windows de escritorio: "Activar o desactivar características de Windows" > Internet Information Services >' -ForegroundColor Yellow
    Write-Host '     Herramientas de administración web > Consola de administración de IIS y Scripts y herramientas.' -ForegroundColor Yellow
    Write-Host '   - Se está ejecutando PowerShell de 32 bits sin acceso al de 64 bits (use el PowerShell x64).' -ForegroundColor Yellow
    Write-Host '   - IIS se instaló o cambió con esta sesión abierta: reinicie el equipo o ejecute "iisreset" y reintente.' -ForegroundColor Yellow
    Abortar 'IIS no está disponible para el instalador. Resuelva lo anterior y vuelva a ejecutarlo.'
}

# ---------- 2. Datos ----------
Titulo 'Datos de instalación'
$sitio = Pedir 'Nombre del sitio IIS' 'PortafolioQA' 'Sitio'
$actualizar = $false
if ($Silencioso -and $P['Modo'] -eq 'actualizar' -and -not (Test-Path "IIS:\Sites\$sitio")) { Abortar "No existe el sitio IIS '$sitio' para actualizar. Use 'Instalación nueva' o corrija el nombre." }
if (Test-Path "IIS:\Sites\$sitio") {
    if ($Silencioso) {
        if ($P['Modo'] -ne 'actualizar') { Abortar "El sitio '$sitio' ya existe. Ejecute de nuevo y elija 'Actualizar una instalación existente'." }
        $r = 'A'
    } else { $r = Pedir "El sitio '$sitio' ya existe. (A)ctualizar archivos, o (C)ancelar" 'A' }
    if ($r -notmatch '^[Aa]') { Abortar 'Cancelado por el usuario.' }
    $actualizar = $true
    $carpeta = (Get-Item "IIS:\Sites\$sitio").physicalPath
    $pool = (Get-Item "IIS:\Sites\$sitio").applicationPool
} else {
    $carpeta = Pedir 'Carpeta del sitio' 'C:\inetpub\portafolio-qa' 'Carpeta'
    $pool = $sitio
}

if (-not $actualizar) {
    $hostSitio = Pedir 'Host del sitio (vacío = cualquier host)' '' 'Host'
    $certs = @(Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) })
    $thumb = $null
    if ($Silencioso) {
        if ($P['CertThumb'] -and $P['CertThumb'] -ne 'AUTOFIRMADO') {
            $thumb = $P['CertThumb'].Trim()
            if (-not (Test-Path "Cert:\LocalMachine\My\$thumb")) { Abortar "El certificado seleccionado ($thumb) ya no está en el almacén del equipo." }
        }
    } elseif ($certs.Count -gt 0) {
        Write-Host 'Certificados disponibles:'
        for ($i = 0; $i -lt $certs.Count; $i++) { Write-Host ('  [{0}] {1}  (vence {2:dd/MM/yyyy})' -f ($i + 1), $certs[$i].Subject, $certs[$i].NotAfter) }
        $n = Pedir 'Número de certificado (0 = crear uno autofirmado de PRUEBA)' '1'
        if ($n -match '^\d+$' -and [int]$n -ge 1 -and [int]$n -le $certs.Count) { $thumb = $certs[[int]$n - 1].Thumbprint }
    } else { Write-Host 'No hay certificados en el almacén del equipo.' -ForegroundColor Yellow }
    if (-not $thumb) {
        $dns = if ($hostSitio) { $hostSitio } else { $env:COMPUTERNAME }
        Write-Host "Se creará un certificado autofirmado para '$dns' (solo para pruebas; los navegadores mostrarán advertencia)." -ForegroundColor Yellow
        $thumb = (New-SelfSignedCertificate -DnsName $dns -CertStoreLocation Cert:\LocalMachine\My -NotAfter (Get-Date).AddYears(2)).Thumbprint
    }
    $puerto = Pedir 'Puerto HTTPS' '443' 'Puerto'
    $claves = Pedir 'Carpeta persistente de claves de protección de datos' 'D:\portafolio-qa-keys' 'CarpetaClaves'
    $unidad = Split-Path $claves -Qualifier
    if (-not (Test-Path "$unidad\")) { $claves = 'C:\portafolio-qa-keys'; Write-Host "La unidad $unidad no existe; se usará $claves" -ForegroundColor Yellow }

    Titulo 'Administrador inicial y Squash TM'
    $admCorreo = Pedir 'Correo del administrador inicial' '' 'AdminCorreo'
    $admPass = PedirSecreto 'Contraseña inicial (se cambia en el primer ingreso)' 'AdminClave'
    if ($Silencioso -and ([string]::IsNullOrWhiteSpace($admCorreo) -or [string]::IsNullOrWhiteSpace($admPass))) { Abortar 'Faltan el correo o la contraseña del administrador inicial.' }
    $squashUrl = Pedir 'URL de Squash TM (vacío = sin integración con Squash)' '' 'SquashUrl'
    $squashToken = PedirSecreto 'Token de API de Squash TM (solo lectura recomendado; Enter para omitir)' 'SquashToken'
    if ($squashToken -and -not $squashUrl) { Abortar 'Para usar un token de Squash TM debe indicar también su URL.' }
}

if ($pool -notmatch '^[\w\.\- ]{1,128}$') { Abortar "El nombre del application pool no es válido: '$pool'." }

# ---------- 4. Archivos y sitio ----------
Titulo 'Publicando archivos'
if ($actualizar) { Stop-WebAppPool $pool -ErrorAction SilentlyContinue; Start-Sleep 3 }
New-Item -ItemType Directory $carpeta -Force | Out-Null
robocopy $RutaApp $carpeta /E /XD logs /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { Abortar 'Falló la copia de archivos.' }
New-Item -ItemType Directory (Join-Path $carpeta 'logs') -Force | Out-Null

if (-not $actualizar) {
    New-Item -ItemType Directory $claves -Force | Out-Null
    if (-not (Test-Path "IIS:\AppPools\$pool")) { New-WebAppPool $pool | Out-Null }
    Set-ItemProperty "IIS:\AppPools\$pool" managedRuntimeVersion ''
    Set-ItemProperty "IIS:\AppPools\$pool" processModel.identityType ApplicationPoolIdentity
    New-Website -Name $sitio -PhysicalPath $carpeta -ApplicationPool $pool -Port 80 -HostHeader $hostSitio -Force | Out-Null
    Remove-WebBinding -Name $sitio -Protocol http -Port 80 -ErrorAction SilentlyContinue
    New-WebBinding -Name $sitio -Protocol https -Port $puerto -HostHeader $hostSitio -SslFlags 0
    (Get-WebBinding -Name $sitio -Protocol https).AddSslCertificate($thumb, 'my')

    $vars = [ordered]@{
        'ASPNETCORE_ENVIRONMENT'                 = 'Production'
        'DataProtection__RutaClaves'             = $claves
        'Seguridad__AdminInicial__Correo'        = $admCorreo
        'Seguridad__AdminInicial__Contrasena'    = $admPass
    }
    if ($squashUrl) { $vars['Squash__BaseUrl'] = $squashUrl }
    if ($squashToken) { $vars['Squash__Token'] = $squashToken }
    foreach ($k in $vars.Keys) {
        $valor = ([string]$vars[$k]).Replace("'", "''")
        & "$env:windir\system32\inetsrv\appcmd.exe" set config -section:system.applicationHost/applicationPools "/+[name='$pool'].environmentVariables.[name='$k',value='$valor']" /commit:apphost | Out-Null
    }
    $idp = "IIS APPPOOL\$pool"
    icacls $carpeta /grant "${idp}:(OI)(CI)RX" /T /Q | Out-Null
    icacls (Join-Path $carpeta 'logs') /grant "${idp}:(OI)(CI)M" /Q | Out-Null
    icacls $claves /grant "${idp}:(OI)(CI)M" /Q | Out-Null
}

# Carpeta de datos del servidor (conexión cifrada, código de configuración y auditoría): solo Administradores, SYSTEM y el pool.
$datos = Join-Path $env:ProgramData 'PortafolioQA'
New-Item -ItemType Directory $datos -Force | Out-Null
$idpDatos = "IIS APPPOOL\$pool"
# Solo se tocan los permisos de la CARPETA (los archivos nuevos los heredan). Primero se conceden y luego se quita la herencia:
# si una concesión falla, la carpeta no queda sin acceso. No se usa /T para no dejar archivos sin permisos heredados.
$resultadoAcl = icacls $datos /grant:r '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-18:(OI)(CI)F' "${idpDatos}:(OI)(CI)M" /Q 2>&1
if ($LASTEXITCODE -ne 0) { Abortar "No se pudieron asignar permisos sobre $datos a '$idpDatos': $($resultadoAcl -join ' ')" }
icacls $datos /inheritance:r /Q | Out-Null
# Repara los archivos que ya existan (p. ej. de una instalación anterior con permisos rotos): toma la propiedad y los
# restablece a los permisos heredados de la carpeta.
if (Get-ChildItem $datos -Force -File -ErrorAction SilentlyContinue) {
    takeown /F (Join-Path $datos '*') /A 2>&1 | Out-Null
    icacls (Join-Path $datos '*') /reset /Q 2>&1 | Out-Null
}
$aclFinal = (icacls $datos) -join ' '
if ($aclFinal -notmatch [regex]::Escape($idpDatos)) { Abortar "La carpeta $datos no quedó con permiso para '$idpDatos'. Concédalo manualmente: icacls ""$datos"" /grant ""${idpDatos}:(OI)(CI)M""" }
Write-Host "OK: la cuenta '$idpDatos' puede modificar $datos." -ForegroundColor Green
Start-WebAppPool $pool -ErrorAction SilentlyContinue
Start-Website $sitio -ErrorAction SilentlyContinue

# ---------- 5. Verificación ----------
if (-not $actualizar) {
    Titulo 'Verificando el sitio'
    $hh = if ($hostSitio) { $hostSitio } else { 'localhost' }
    $baseUrl = "https://${hh}:$puerto"
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $estado = ''
    for ($i = 0; $i -lt 15 -and $estado -notin @('ok', 'pendiente'); $i++) {
        Start-Sleep 2
        try { $estado = (Invoke-WebRequest "$baseUrl/healthz" -UseBasicParsing -TimeoutSec 10).Content.Trim() }
        catch {
            # /healthz responde 503 con "pendiente" mientras la conexión a la base no esté configurada: es lo esperado.
            $estado = ''
            if ($_.Exception.Response) { try { $estado = (New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd().Trim() } catch {} }
        }
    }
    if ($estado -eq 'ok') { Write-Host "OK: $baseUrl responde (la aplicación ya está activa)." -ForegroundColor Green }
    elseif ($estado -eq 'pendiente') { Write-Host "OK: el sitio responde y está pendiente de configurar la conexión a la base de datos." -ForegroundColor Green }
    else { $codigoSalida = 2; Write-Host "ATENCIÓN: $baseUrl/healthz no respondió. Revise logs\ (stdoutLogEnabled en web.config)." -ForegroundColor Yellow }
}
Write-Host ''
Write-Host 'Instalación terminada.' -ForegroundColor Green
if (-not $actualizar) {
    Write-Host 'La aplicación quedó instalada pero SIN ACTIVAR: falta configurar la conexión a la base de datos.' -ForegroundColor Yellow
    Write-Host "  1. Abra $baseUrl/Setup" -ForegroundColor Yellow
    Write-Host "  2. Escriba el código de configuración que está en: $(Join-Path $datos 'codigo-configuracion.txt')" -ForegroundColor Yellow
    Write-Host '  3. Indique servidor, base y autenticación, pruebe la conexión y active la aplicación.' -ForegroundColor Yellow
    Write-Host "  4. Ya activa, ingrese con el administrador inicial, cambie su contraseña y elimine las variables Seguridad__AdminInicial__* del pool '$pool'." -ForegroundColor Yellow
} else {
    Write-Host 'Archivos actualizados. La conexión a la base de datos no se modificó.' -ForegroundColor Yellow
    Write-Host '  Si la versión nueva trae cambios de base de datos, un Administrador verá el aviso y debe aplicarlos en' -ForegroundColor Yellow
    Write-Host '  Administración > Conexión a la base de datos (aplicar migraciones o ejecutar el script que ofrece el asistente).' -ForegroundColor Yellow
}
Terminar $codigoSalida
