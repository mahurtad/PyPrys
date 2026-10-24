# Ficha d80a1a1 / servidor IIS en uso

| Dato | Valor |
|---|---|
| Estado | En mantenimiento: 2 migraciones pendientes, reportado por el usuario |
| Fecha del reporte | 19/10/2026, America/Lima |
| Servidor / URL / responsable | Por confirmar |
| Commit instalado anterior | Desconocido; no asumir que es el anterior de master |
| Commit objetivo documental | d80a1a17b8a43fc502d27d4590f11d0a200f0dab; falta confirmar correspondencia del paquete instalado |
| Paquete / SHA-256 | Pendiente de registrar |
| Base / servidor SQL / conexión efectiva | Pendiente de confirmar sin registrar secretos |
| Migraciones pendientes | Cantidad 2; nombres aún no confirmados |
| Requisitos | [Versión y escenarios](version-d80a1a1.md) |
| Squash / Bugzilla / motores / URL | Por inventariar |
| Respaldo y recuperación | Pendiente de registrar antes de aplicar migraciones |

Mensaje reportado: «Faltan 2 migraciones por aplicar a la base de datos. La aplicación permanece en mantenimiento hasta aplicar las migraciones.»

Dos migraciones pendientes no permiten deducir el commit anterior instalado.
Podrían ser Etapa17 y Etapa18 si la base está en Etapa16, pero deben confirmarse los nombres desde la pantalla o comparando el historial SQL con Migrations.

| Paso | Responsable / fecha / resultado / evidencia |
|---|---|
| Actualización de archivos | Usuario reporta actualización realizada; faltan paquete, responsable y log |
| Confirmar base efectiva y nombres de las 2 pendientes | Pendiente |
| Respaldar base, configuración y claves | Pendiente |
| Aplicar pendientes mediante administrador autorizado o DBA | Pendiente |
| Confirmar historial hasta Etapa18 y sin pendientes | Pendiente |
| Configuración de integraciones, incluida AdminApiKey si falta habilitar productos automáticos | Pendiente |
| Actualizar desde Squash y verificar módulos | Pendiente |
| Verificar Bugzilla y asociación Squash-producto | Pendiente |
| Health 200/ok, login y Dashboard | Pendiente |
| Cierre validado y archivo de evidencias | Pendiente |

No se accedió al servidor ni se aplicaron migraciones al crear esta ficha.

## Incidencia SQL reportada el 19/10/2026

El usuario reporta Msg 207, Invalid column name NombreCompleto, líneas 643–647 del script completo, servidor SVRCERSSISMO, base PortafolioQA. La captura también muestra dos filas afectadas; no acredita qué migraciones terminaron.
Causa localizada: UPDATE de Etapa6 en SQL idempotente referencia una columna eliminada en bases que ya tienen esa migración. SQL Server puede validarlo aunque el IF omita su ejecución.
Corrección local: diferir compilación de los UPDATE de Etapa6 mediante EXEC(N'...'). No requiere una nueva migración ni recrear columnas.
Pendiente en servidor: consultar historial completo y nombres de pendientes antes de reintentar; no insertar filas manualmente en __EFMigrationsHistory. No se ejecutaron cambios en el servidor desde esta revisión.

## SQL corregido para ejecución manual

Archivo: [PortafolioQA-Servidor-Corregido.sql](PortafolioQA-Servidor-Corregido.sql).
Generado a partir del SQL del paquete con compilación diferida de los UPDATE históricos de Etapa6 y opciones SET para índices.
Validado el 19/10/2026 en base temporal independiente sobre SQL Server 2019 Express: instalación completa de 18 migraciones y segunda ejecución idempotente, ambas correctas. La base temporal fue eliminada tras la prueba; no se modificó la base habitual ni el servidor.

En servidor: respaldar, seleccionar la base PortafolioQA en SSMS y ejecutar el archivo completo sin seleccionar fragmentos. Después consultar __EFMigrationsHistory, confirmar ausencia de pendientes y comprobar /healthz. Si sigue en mantenimiento, reciclar el pool de Portafolio y volver a comprobar.
