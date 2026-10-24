# Restauración local — 23/10/2026

Equipo: `LT-LTP8119HG-MH`. Proyecto en `G:\My Drive\Data Analysis\PyPrys\Proyectos\Laureate\portafolio-qa`.

## Configuración recuperada

- SQL Server: `LT-LTP8119HG-MH\SQLEXPRESS`, base `PortafolioQA`, autenticación Windows.
- La base restaurada contiene 18 migraciones, todas coincidentes con el código; no se aplicaron migraciones adicionales.
- Datos de Portafolio: 14 proyectos y 3 usuarios.
- SDK .NET 8.0.426; ASP.NET Core 8.0.32. Compilación correcta con una advertencia CS8602 preexistente.
- Se restauraron los seis volúmenes del respaldo `PortafolioQA_20261023_111345`, verificando previamente sus SHA-256.
- Squash conserva 4 proyectos; Bugzilla conserva 4 productos y 3 defectos.
- Se recuperaron los user-secrets en el perfil actual y se actualizó únicamente su conexión a SQL Server.
- Las claves antiguas de Data Protection usan DPAPI; no se importaron. El nuevo perfil genera sus claves locales, y el inicio de sesión fue verificado.

## Accesos y operación

| Servicio | Dirección |
| --- | --- |
| Portafolio | http://localhost:5011 |
| Salud de Portafolio y SQL | http://localhost:5011/healthz |
| Squash TM | http://localhost:8090/squash |
| Bugzilla | http://localhost:8080 |

Con Docker Desktop iniciado, abrir una nueva consola de PowerShell en la raíz del proyecto y ejecutar:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Iniciar-Todo.ps1
```

La política de este equipo impide ejecutar scripts por defecto; el comando permite únicamente esa ejecución y no cambia la política persistente. Portafolio permanece asociado a esa consola. Para detenerlo, usar Ctrl+C. Para detener los servicios Docker conservando los datos:

```powershell
docker compose -f infra/local/compose.yaml stop
```

No usar `down -v`: elimina los volúmenes con datos.

Las credenciales de las integraciones siguen en `infra/local/credentials.local.txt` y los secretos de Portafolio en `%APPDATA%\Microsoft\UserSecrets\29d62801-1fe2-4114-8818-ad1d23f5787d\secrets.json`. No deben versionarse.

## Ajustes de este equipo y validación

Docker no pudo montar el archivo de plugins directamente desde la unidad virtual G:. Se copió a `%LOCALAPPDATA%\PortafolioQA\start-plugins.cfg`, y `.env` define `SQUASH_PLUGINS_CONFIG` con esa ruta. Compose permite esta configuración y mantiene la ruta relativa por defecto en otros equipos.

`infra/local/Verificar.ps1` ahora admite Windows PowerShell 5.1 y envía la clave de Bugzilla como parámetro, igual que el cliente de la aplicación. Los fallos de esa consulta no imprimen la URL con la clave.

Validado: `/healthz` devuelve `200 ok`; inicio de sesión y páginas Portafolio, Gantt, Calendario, Capacidad y Administración de proyectos responden HTTP 200. Las dos claves de Bugzilla permiten consultas autenticadas. Squash acepta el token restaurado y devuelve los 4 proyectos. El verificador completo termina correctamente y las 64 pruebas automatizadas pasan.

Para repetir la comprobación de las integraciones:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\infra\local\Verificar.ps1
```
