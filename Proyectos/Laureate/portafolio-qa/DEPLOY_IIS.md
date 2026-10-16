﻿# Despliegue en IIS (Windows Server)

Guía para publicar Portafolio QA como versión interna. Nada de esto se ejecuta
automáticamente: son pasos para quien administre el servidor.

## 1. Requisitos del servidor

- Windows Server con el rol **Web Server (IIS)**.
- **.NET 8 Hosting Bundle** (incluye el ASP.NET Core Module V2). Reiniciar IIS tras instalarlo (`iisreset`).
- Acceso a la instancia de SQL Server de destino.
- Un certificado HTTPS para el sitio. La cookie de sesión solo viaja por HTTPS fuera de Development.

## 2. Publicar

```powershell
dotnet publish -c Release -o .\publish
```

Copiar el contenido de `publish\` a la carpeta del sitio (p. ej. `C:\inetpub\portafolio-qa`).
`publish` genera el `web.config` (hosting *in-process*). No incluye secretos.

## 3. Base de datos

La instalación **no toca SQL Server**. La conexión se configura después, desde el asistente de la aplicación
(`https://<servidor>/Setup`, con el código de `C:\ProgramData\PortafolioQA\codigo-configuracion.txt`). El asistente prueba
conexión, permisos y esquema, genera el script de acceso para el DBA y activa la aplicación (ver `Instalador\LEEME.txt`).

Si se prefiere preparar la base a mano, el script de esquema se obtiene del asistente o con:

```powershell
dotnet ef migrations script --idempotent --configuration Release -o portafolio-qa.sql
```

La cuenta de la app necesita `db_datareader` y `db_datawriter` (no `db_owner`). Para aplicar migraciones desde el asistente,
además `db_ddladmin`; si no se otorga, el DBA ejecuta el script de esquema.

## 4. Sitio y application pool

1. Application pool nuevo: **.NET CLR version = Sin código administrado**, identidad `ApplicationPoolIdentity`
   (o una cuenta de servicio dedicada).
2. Sitio apuntando a la carpeta publicada, con binding HTTPS.
3. Permisos de la carpeta: lectura/ejecución a la identidad del pool; escritura solo en `logs\` y en la
   carpeta de claves (paso 6).
4. Si el acceso a SQL es con autenticación integrada, esa identidad (`IIS APPPOOL\<nombre>` o la cuenta de
   servicio) necesita un login en SQL Server.

## 5. Configuración y secretos

Todo se define como **variables de entorno del application pool** (o en el `web.config` del servidor,
fuera del repositorio). Doble guion bajo = separador de nivel:

| Variable | Valor |
|---|---|
| `ASPNETCORE_ENVIRONMENT` | `Production` |
| `ConnectionStrings__PortafolioQAContext` | **Opcional/heredada.** Las instalaciones nuevas no la definen: la conexión se guarda cifrada desde `/Setup`. Si existe y no hay `conexion.dat`, se usa |
| `DataProtection__RutaClaves` | carpeta persistente, p. ej. `D:\portafolio-qa-keys` |
| `Seguridad__AdminInicial__Correo` | correo del primer administrador (solo primer arranque) |
| `Seguridad__AdminInicial__Contrasena` | contraseña inicial (solo primer arranque) |

- El administrador inicial se crea **solo si no existe ningún usuario**. Tras el primer ingreso, cambia la
  contraseña desde la app y **elimina las dos variables `Seguridad__AdminInicial__*`**.
- Sin `DataProtection__RutaClaves`, cada reciclaje del pool invalida las sesiones y los tokens antiforgery.
  La carpeta debe existir y ser escribible por la identidad del pool, y conviene incluirla en el respaldo.

## 6. Verificación tras publicar

- `https://<servidor>/healthz` responde `ok` (200); 503 si no hay conexión a la base.
- Abrir el sitio: redirige a `/Account/Login`. Entrar con el administrador inicial.
- Crear un usuario *Lector* y confirmar que ve Portafolio y Planificación pero no Administración.
- Revisar `logs\` (activar `stdoutLogEnabled="true"` en `web.config` solo para diagnosticar y volver a apagarlo).

## 7. Roles

| Rol | Puede |
|---|---|
| Administrador | Todo, incluida la gestión de usuarios (Administración → Usuarios) |
| Editor | Consultar y modificar datos (proyectos, asignaciones, catálogos, hitos, feriados, ausencias) |
| Lector | Solo consultar Portafolio, Gantt, Calendario y Capacidad |

## 8. Pendientes conscientes

- No hay cabecera `Content-Security-Policy`: las páginas usan estilos y manejadores en línea. Evaluar al
  endurecer el sitio.
- No hay recuperación de contraseña por correo: un administrador restablece la clave desde Usuarios.
- No hay registro de auditoría de cambios de datos (quién modificó qué).
- Bugzilla sigue sin conexión real (etapa 5).
