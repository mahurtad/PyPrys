# Portafolio QA

Aplicación web interna (ASP.NET Core + Razor Pages) para administrar
iniciativas QA, planificar recursos y, en etapas posteriores, consultar
certificación desde Squash TM y defectos desde Bugzilla.

Consulta [INSTRUCTIVO.md](INSTRUCTIVO.md) para el alcance completo. Estado:
**etapa 1** (prototipo: tarjetas, tabla, filtros, detalle y administración) y
**etapa 2** (planificación: Gantt, calendario y capacidad) y **etapa 3**
(autenticación, roles y despliegue), con clientes reales para Squash TM y Bugzilla.

## Squash TM y Bugzilla locales

La configuración actual usa **http://localhost:8090/squash** y **http://localhost:8080**.
La instalación, el inicio y las asociaciones se describen en [infra/local/README.md](infra/local/README.md).
Squash TM dispone de cliente REST y sincronización de campañas, iteraciones y casos.
Bugzilla dispone de cliente REST y consulta de defectos por producto en el detalle de la iniciativa.
Las credenciales se almacenan en `dotnet user-secrets`; SQL Server sigue siendo la base del portafolio.
La creación automática del producto de Bugzilla por iniciativa (permisos, pasos del DBA y limitaciones) está en [INTEGRACION_BUGZILLA.md](INTEGRACION_BUGZILLA.md).
Las secciones de etapas que siguen describen el alcance histórico del prototipo.

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
6. Desde **Administración → Proyectos** vincula el proyecto de Squash TM y
   sincroniza su ejecución. Desde **Administración → Integraciones** asocia
   Bugzilla usando el nombre exacto del producto para consultar sus defectos.

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
| `Services/Integrations` | Clientes REST de Squash TM y Bugzilla; sin credenciales muestran integración no configurada |
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

## Integraciones disponibles

- Squash TM: cliente REST, vínculo de proyecto y sincronización de ejecución.
- Bugzilla: consulta REST de defectos por producto y enlaces desde el detalle.


## Dashboard gerencial (`/Dashboard`)

Lectura por proyecto que integra Portafolio, Squash TM y Bugzilla. **Solo lee**: usa los datos de Squash ya sincronizados (no llama a Squash ni
sincroniza) y consulta Bugzilla por GET con caché de 60 s (15 s si falla) y como máximo 4 consultas simultáneas, una por producto distinto.
Todos los roles autenticados lo ven (Lector incluido). Código: `Services/Dashboard/` y `Pages/Dashboard/`.

**Presentación**: usa el mismo `_Layout` y menú de Tabler 1.6.1 que el resto de la aplicación (sin layout aparte, sin AdminLTE ni Bootstrap duplicado;
el CSS propio es `wwwroot/css/dashboard.css`, cargado solo aquí mediante la sección `Styles`). Los gráficos (barras apiladas y dona) son HTML/CSS
renderizado en servidor: sin biblioteca de gráficos ni JavaScript y sin licencias adicionales. Cada gráfico lleva leyenda, cifras escritas,
`aria-label` y un resumen en texto para lectores de pantalla. Capturas de referencia con datos de demostración en `docs/dashboard/`.

**Alcance** (el mismo en indicadores, gráficos por estado y tipo, tabla y alertas): todos los proyectos del catálogo de la institución elegida (o de
todas), activos o no y de cualquier estado, como el Portafolio. Los gráficos de pruebas y defectos muestran solo los proyectos con cifras de esa fuente
y listan debajo el resto con su motivo. Las tarjetas por institución cuentan siempre el portafolio completo; al elegir una se filtra todo lo demás.
«Con defectos bloqueantes» solo se evalúa donde Bugzilla respondió.

**Fórmulas** (por proyecto; nunca se suman ni promedian entre proyectos; no hay semáforo ni índice global):

- Avance del proyecto: valor guardado en Portafolio (con Squash, promedio del peso de cada caso único; sin Squash, actividades ponderadas). No es % ejecutado.
- Casos totales: casos de prueba únicos y vigentes que cuentan (excluye No ejecutable, Omitido, Cancelado, No encontrado, que se informan aparte).
- Ejecutados = aprobados (SUCCESS, SETTLED) + fallidos (FAILURE, ERROR) + advertencia (WARNING). Bloqueados = BLOCKED (no ejecutados). En curso = RUNNING; sin ejecutar = el resto (READY, NOT_RUN, desconocidos). Los seis estados suman los casos totales.
- % ejecutado = ejecutados / casos. Tasa de aprobación = aprobados / ejecutados (sin ejecuciones: «Sin ejecuciones», no 0 %).
- Defectos abiertos = estado distinto de RESOLVED, VERIFIED y CLOSED; bloqueantes / no bloqueantes / mejoras / sin clasificar según severidad (reglas de `DetalleModel.ResumenDefectos`); pendientes de retest = RESOLVED + FIXED; cada defecto cuenta una vez por número.
- Atención requerida (hechos): fecha comprometida vencida con proyecto no Cerrado ni Cancelado; casos bloqueados > 0; defectos bloqueantes abiertos > 0; pendientes de retest > 0.

**Estados de fuente**: Con datos (incluye ceros reales), Sin datos, Sin vínculo, Integración no configurada, Error. Una fuente con error no genera cifras ni hechos y el resto sigue.
Bugzilla responde lista vacía ante un producto inexistente o sin permiso; con 0 defectos se verifica que el producto exista (`product?names=`) y, si no, se marca Error.

**Filtros**: institución (vista general) y proyecto (abre el resumen). Campaña, iteración y módulo (suite de Squash) solo afectan a las pruebas, por ID y por la
jerarquía real sincronizada; Bugzilla no informa esa relación, por lo que los defectos son siempre del producto completo.
La API de Squash expone `issues` por campaña, iteración y suite, pero no se usa: exigiría consultas por iteración en cada visita y confirmar que el bugtracker
de Squash es el mismo Bugzilla del producto.
