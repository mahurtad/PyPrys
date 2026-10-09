# Portafolio QA

Aplicación web interna (ASP.NET Core + Razor Pages) para administrar
iniciativas QA, planificar recursos y, en etapas posteriores, consultar
certificación desde Squash TM y defectos desde Bugzilla.

Consulta [INSTRUCTIVO.md](INSTRUCTIVO.md) para el alcance completo. Estado:
**etapa 1** (prototipo: tarjetas, tabla, filtros, detalle y administración) y
**etapa 2** (planificación: Gantt, calendario y capacidad) y **etapa 3**
(autenticación, roles y despliegue). No incluye conexiones reales a Squash TM/Bugzilla.

## Tecnología

- .NET 8 (LTS) — ASP.NET Core con Razor Pages
- Entity Framework Core 8 + SQL Server
- Bootstrap 5 + JavaScript (incluidos en el proyecto)

## Requisitos previos

- [SDK de .NET 8](https://dotnet.microsoft.com/download/dotnet/8.0)
- SQL Server o [SQL Server LocalDB](https://learn.microsoft.com/sql/database-engine/configure-windows/sql-server-express-localdb)
  (LocalDB viene con Visual Studio y es suficiente para desarrollo local)
- Herramienta `dotnet-ef` (se instala con el paso 2 abajo)

## Configuración

La cadena de conexión vive en `appsettings.json` bajo `ConnectionStrings:PortafolioQAContext`
y apunta a LocalDB por defecto:

```json
"ConnectionStrings": {
  "PortafolioQAContext": "Server=(localdb)\\mssqllocaldb;Database=PortafolioQA;Trusted_Connection=True;MultipleActiveResultSets=true;TrustServerCertificate=True"
}
```

Para un SQL Server real, **no edites `appsettings.json` con credenciales**.
Sobrescribe la cadena de conexión mediante una variable de entorno o
`dotnet user-secrets` en desarrollo:

```powershell
dotnet user-secrets init
dotnet user-secrets set "ConnectionStrings:PortafolioQAContext" "Server=mi-servidor;Database=PortafolioQA;User Id=...;Password=...;TrustServerCertificate=True"
```

o en producción, mediante la variable de entorno `ConnectionStrings__PortafolioQAContext`.

## Pasos para ejecutar

```powershell
# 1. Restaurar paquetes
dotnet restore

# 2. Instalar la herramienta de EF Core (una sola vez por máquina)
dotnet tool install --global dotnet-ef

# 3. Crear la base de datos aplicando las migraciones
dotnet ef database update

# 4. Ejecutar la aplicación
dotnet run
```

La aplicación abrirá en `https://localhost:xxxx` (el puerto exacto se
muestra en la consola). La página inicial redirige al Portafolio.

## Primer uso

El catálogo de instituciones **comienza vacío a propósito** (ver
INSTRUCTIVO.md, sección 4). Para ver datos en el portafolio:

1. Ve a **Administración → Instituciones** y registra al menos una.
2. Ve a **Administración → Recursos QA** y registra al menos uno.
3. (Opcional) Registra **Sistemas / periféricos** y relaciónalos con instituciones.
4. Ve a **Administración → Proyectos** y crea una iniciativa, eligiendo
   institución, tipo (`Proyecto`, `Requerimiento no BAU` o `BAU`), estado,
   fechas y sistemas impactados.
5. Desde el detalle del proyecto (o desde **Administración → Asignaciones**)
   agrega recursos QA al proyecto con fechas, rol y porcentaje de dedicación.
6. (Opcional) Desde **Administración → Integraciones** asocia el proyecto con
   un identificador externo de Squash TM o Bugzilla. En esta etapa solo se
   guarda la asociación; no hay sincronización real todavía.

Todos los cambios persisten en SQL Server y se conservan al recargar.

## Etapa 2: planificación y capacidad

Menú **Planificación**:

- **Gantt** (`/Planificacion/Gantt`): vista *por proyecto* (barra del proyecto con
  relleno = avance, asignaciones QA debajo, hitos) y *por recurso QA* (franja de
  carga diaria + sus asignaciones). Filtros de periodo e institución.
- **Calendario** (`/Planificacion/Calendario`): mes con periodos de asignación,
  ausencias, hitos y feriados; filtros por recurso y proyecto.
- **Capacidad QA** (`/Planificacion/Capacidad`): por periodo y equipo, carga
  promedio y pico, disponibilidad, días sobreasignados, horas asignadas/capacidad
  y los rangos de fechas sobreasignados con los proyectos involucrados.

Menú **Administración** (nuevo): **Hitos**, **Feriados** y **Ausencias**.

Gantt y calendario son componentes propios (HTML/CSS renderizado en servidor, sin
librerías de terceros), por lo que no hay licencias que revisar. Si más adelante
se necesita arrastrar y soltar, habrá que elegir una librería y validar su licencia.

### Reglas de cálculo (`Services/CapacidadService.cs`)

- Día laborable = lunes a viernes que no sea feriado registrado. En días no
  laborables no hay capacidad ni carga.
- Carga diaria = suma de los porcentajes de las asignaciones vigentes ese día.
  Capacidad 100% + dos asignaciones de 60% = carga 120%.
- Disponibilidad = máx(0, capacidad - carga); el exceso (carga - capacidad) se muestra aparte.
- Una ausencia deja la capacidad del día en 0: si hay una asignación vigente, cuenta
  como sobreasignación (conflicto a resolver).
- Las iniciativas **canceladas** no consumen capacidad; las cerradas sí cuentan en
  su periodo histórico. Solo se consideran recursos QA activos.
- Horas = porcentaje x horas por día laborable. Configurable en `appsettings.json`
  (`Planificacion:HorasPorDia`, 8 por defecto).
- Periodo máximo de consulta: 400 días.
- Al crear o editar una asignación se avisa si el recurso queda sobreasignado
  (no se bloquea el guardado).
- Los feriados **no vienen precargados**: registra los oficiales en Administración → Feriados.

### Aplicar la migración

La etapa 2 agrega tres tablas (`Feriados`, `Ausencias`, `Hitos`):

```powershell
dotnet ef database update
```

## Etapa 3: acceso, roles y despliegue

Toda la aplicación exige iniciar sesión (cuentas locales con ASP.NET Core Identity).

| Rol | Puede |
|---|---|
| Administrador | Todo, incluida **Administración → Usuarios** |
| Editor | Consultar y modificar datos |
| Lector | Solo consultar Portafolio, Gantt, Calendario y Capacidad |

- **Primer administrador:** se crea al arrancar, solo si no hay usuarios, con
  `Seguridad:AdminInicial:Correo` y `Seguridad:AdminInicial:Contrasena`. En desarrollo:

  ```powershell
  dotnet user-secrets set "Seguridad:AdminInicial:Correo" "tu.correo@dominio"
  dotnet user-secrets set "Seguridad:AdminInicial:Contrasena" "una-clave-de-10+-caracteres-con-Mayus-y-1-numero"
  dotnet ef database update
  dotnet run
  ```
- Contraseñas: mínimo 10 caracteres con mayúscula, minúscula y número. Bloqueo de 15 min tras 5 intentos fallidos.
- Los usuarios no se eliminan: se desactivan (pierden la sesión en ~1 minuto). Siempre debe quedar un
  administrador activo; nadie puede desactivarse a sí mismo.
- Cierre de sesión solo por POST; formularios con token antiforgery; `/healthz` anónimo para monitoreo.
- Despliegue en IIS, variables de entorno y secretos: ver [DEPLOY_IIS.md](DEPLOY_IIS.md).

## Estructura del proyecto

| Carpeta | Contenido |
|---|---|
| `Pages/Portafolio` | Listado (tarjetas/tabla con filtros) y detalle de iniciativas |
| `Pages/Admin` | CRUD de Instituciones, Recursos QA, Sistemas, Proyectos, Asignaciones e Integraciones; altas de Hitos, Feriados y Ausencias |
| `Pages/Planificacion` | Gantt, calendario y capacidad |
| `Pages/Account` | Inicio/cierre de sesión, cambio de contraseña, acceso denegado |
| `Services` | `CapacidadService` / `CalculadoraCapacidad` (cálculo de carga y disponibilidad) |
| `Models` | Entidades y enumeraciones del dominio |
| `Data` | `PortafolioQAContext` (EF Core) |
| `Services/Integrations` | Contratos para Squash TM y Bugzilla; implementaciones "no configuradas" hasta las etapas 4 y 5 |
| `wwwroot` | Bootstrap, jQuery y JavaScript propio (alternancia tarjetas/tabla) |

## Reglas de negocio ya implementadas

- Las instituciones, recursos QA, sistemas y proyectos no se eliminan: se
  desactivan para conservar su historial. Los registros inactivos no
  aparecen en selectores de nuevas asociaciones.
- Se valida que la fecha de inicio sea anterior o igual a la fecha de fin
  en proyectos y asignaciones.
- El porcentaje de dedicación de una asignación debe estar entre 0 y 100.
- Tarjetas y tabla del portafolio comparten los mismos filtros y el mismo
  conjunto de datos.

## Pendiente para próximas etapas

- Etapa 4: cliente real de Squash TM (reemplaza `IntegracionCertificacionNoConfigurada`).
- Etapa 5: cliente real de Bugzilla (reemplaza `IntegracionDefectosNoConfigurada`).
