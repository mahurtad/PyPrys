<#
 Utilidades del asistente gráfico (Setup.exe). Se ejecuta oculto antes de mostrar el asistente.
   -Accion Certificados : escribe en -Salida una línea "huella|descripción" por certificado válido del equipo.
 Compatible con Windows PowerShell 5.1.
#>
param(
    [Parameter(Mandatory)][ValidateSet('Certificados')][string]$Accion,
    [Parameter(Mandatory)][string]$Salida
)
$ErrorActionPreference = 'Stop'

$lineas = @(Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
    Where-Object { $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
    ForEach-Object { '{0}|{1}  (vence {2:dd/MM/yyyy})' -f $_.Thumbprint, $_.Subject, $_.NotAfter })
[IO.File]::WriteAllLines($Salida, [string[]]$lineas, (New-Object System.Text.UTF8Encoding($true)))
exit 0
