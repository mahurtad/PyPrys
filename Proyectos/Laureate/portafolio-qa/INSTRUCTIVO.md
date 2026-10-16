# Instructivo de desarrollo — Portafolio QA

## 1. Objetivo

Crear una aplicación web interna para administrar iniciativas QA, planificar recursos y consultar el avance de certificación desde Squash TM. En una etapa posterior, incorporar defectos desde Bugzilla.

Este documento define el alcance y orienta el desarrollo. No implica construir ni publicar la aplicación todavía.

## 2. Organización sencilla

Mantener todo el proyecto en una carpeta principal llamada `portafolio-qa`, con subcarpetas por responsabilidad. Usar un único proyecto ASP.NET Core con Razor Pages, organizado por responsabilidad.

| Ruta propuesta | Contenido |
|---|---|
| `INSTRUCTIVO.md` | Alcance y reglas de desarrollo |
| `README.md` | Pasos para instalar, configurar y ejecutar |
| `Pages/` | Pantallas y formularios Razor Pages |
| `wwwroot/` | Estilos, JavaScript y recursos visuales |
| `Models/` y `Data/` | Entidades y acceso a SQL Server |
| `Services/` | Lógica de capacidad y servicios |
| `Services/Integrations/` | Conectores Squash TM y Bugzilla |
| `appsettings.json` | Configuración sin credenciales reales |
| `.gitignore` | Exclusión de secretos y archivos temporales |

Una sola carpeta facilita trabajar con Codex y mantener el proyecto. No significa colocar todo el código en un único archivo. Los datos operativos deben persistir en SQL Server. Guardar secretos en variables de entorno o en un gestor de secretos.

### Tecnología definida

- **Web:** ASP.NET Core con Razor Pages, Bootstrap y JavaScript.
- **Datos:** SQL Server con Entity Framework Core.
- **Integraciones:** consultas a las APIs desde el servidor mediante HttpClient.
- **Alojamiento previsto:** IIS en Windows Server.
- **Gantt y calendario:** componentes JavaScript; revisar su licencia antes de elegirlos.

Seleccionar una versión LTS de .NET con soporte vigente al iniciar el desarrollo. Mantener interfaz y backend en el mismo proyecto.

## 3. Módulos y vistas

### Portafolio

- Listar iniciativas en tarjetas y tabla, con los mismos datos y filtros.
- Clasificar mediante un campo único: `Proyecto`, `Requerimiento no BAU` o `BAU`. Validar esta nomenclatura con el equipo antes de cargar datos reales.
- Filtrar por institución, tipo, estado, recurso QA, PM, líder y sistema impactado.
- Seleccionar instituciones desde el catálogo registrado en la aplicación, sin nombres predefinidos.
- Abrir un detalle con resumen, equipo, planificación, sistemas y certificación.

### Planificación y capacidad

- Gantt con vista por proyecto y por recurso QA.
- Calendario con periodos de asignación e hitos.
- Registrar porcentaje de dedicación y fechas por asignación.
- Mostrar carga, disponibilidad y sobreasignación por persona y fecha.
- Permitir varios QA por iniciativa y varias iniciativas por QA.

### Dashboards

- Resumen del portafolio: iniciativas por tipo, institución y estado.
- Capacidad QA: personas disponibles y sobreasignadas en el periodo seleccionado.
- Certificación: avance y resultados provenientes de Squash TM.
- Mantener separados el avance del proyecto, el avance de ejecución de pruebas y los resultados de pruebas.

### Administración

- Instituciones: registrar, editar y desactivar desde la aplicación, con nombre y estado activo. El catálogo comienza vacío y se alimenta mediante este panel.
- Recursos QA: nombres, apellidos, correo, equipo Change o RUN/BAU, capacidad disponible y estado activo.
- Proyectos: nombre, descripción breve, institución, tipo, estado, PM, líder, inicio, fin, avance y sistemas impactados.
- Asignaciones: proyecto, recurso QA, inicio, fin, porcentaje y rol.
- Sistemas/periféricos: nombre, instituciones relacionadas y estado activo.
- Integraciones: asociación de proyectos con sus fuentes externas y estado de sincronización.

Un recurso QA no necesita necesariamente una cuenta de acceso. Gestionar usuarios y permisos de la aplicación por separado. Para la versión compartida, contemplar como mínimo administrador, editor y lector.

## 4. Datos y reglas básicas

| Entidad | Relaciones principales |
|---|---|
| Institución | Proyectos y sistemas/periféricos relacionados |
| Proyecto | Institución, asignaciones y sistemas impactados |
| Recurso QA | Varias asignaciones |
| Asignación | Un proyecto y un recurso, con fechas y dedicación |
| Sistema/periférico | Una o varias instituciones y varios proyectos |
| Asociación externa | Proyecto local, herramienta e identificadores externos |

- Usar identificadores internos estables; no relacionar registros solo por nombre.
- Validar que inicio sea anterior o igual a fin y que cada asignación esté entre 0% y 100%.
- Calcular la carga diaria sumando las asignaciones vigentes ese día. Una persona con capacidad del 100% y dos asignaciones simultáneas de 60% tendrá una carga del 120%.
- Comparar la carga con la capacidad disponible de la persona. Mostrar disponibilidad como máximo entre cero y capacidad menos carga; mostrar el exceso aparte.
- Definir días laborables, feriados y ausencias antes de convertir porcentajes a horas.
- No confundir dedicación con avance: estar asignado al 50% no significa haber completado el 50% del proyecto.
- Evitar nombres duplicados de instituciones y usar el catálogo en formularios, filtros y dashboards. No fijar instituciones en el código ni precargarlas en los datos de demostración.
- Evitar eliminar instituciones, recursos o sistemas que ya tengan historial; permitir desactivarlos conservando sus relaciones.

## 5. Integración Squash TM

Squash TM seguirá administrando casos, campañas, iteraciones y ejecuciones. La web consultará y consolidará esos datos, inicialmente en modo lectura.

1. Confirmar versión instalada, API disponible, autenticación, permisos y conectividad desde el servidor.
2. Asociar cada proyecto local con uno o varios proyectos de Squash y delimitar las campañas o iteraciones incluidas.
3. Consultar datos desde el backend y normalizar los estados realmente disponibles.
4. Mostrar totales, pendientes, aprobados, fallidos, bloqueados y avance por campaña/iteración, según lo que permita la API.
5. Mostrar última actualización, opción de actualizar y errores de conexión. Conservar el último dato válido identificándolo como desactualizado.

Antes de calcular indicadores, acordar qué estados cuentan como ejecutados. Usar elementos del plan de pruebas dentro del alcance seleccionado como denominador, no todos los casos del repositorio. Para resultados actuales, tomar el último resultado de cada elemento y evitar contar reejecuciones como casos nuevos.

**Avance de ejecución = elementos ejecutados / elementos planificados × 100.** Si no hay elementos planificados, mostrar “Sin plan de pruebas”; no presentar un porcentaje ficticio. La aprobación se calcula por separado y ejecutar todas las pruebas no equivale a obtener la certificación.

Los endpoints y campos concretos se definirán al revisar la API de la instalación; no asumir que están disponibles.

## 6. Defectos con Bugzilla — etapa posterior

- Mantener Bugzilla como gestor de defectos y comenzar con consultas de lectura.
- Asociar proyectos locales con productos, componentes u otro criterio explícito de Bugzilla.
- Mostrar ID, título, estado, resolución, severidad, prioridad, responsable y enlace al defecto.
- Incorporar totales y abiertos/cerrados según los estados configurados en Bugzilla.
- Relacionar defectos con ejecuciones o casos solo cuando exista un vínculo verificable en las fuentes. No inferir relaciones por títulos parecidos.
- Calcular tiempos de resolución únicamente si se dispone del historial necesario.

## 7. Implementación por etapas

| Etapa | Entregable |
|---|---|
| 1. Prototipo | Tarjetas, tabla, filtros, detalle y administración con datos de demostración |
| 2. Planificación | Asignaciones, Gantt, calendario y cálculo de capacidad |
| 3. Versión interna | Persistencia, autenticación, permisos y configuración de despliegue |
| 4. Squash TM | Conexión real, asociaciones y dashboards de certificación |
| 5. Bugzilla | Defectos por proyecto e indicadores asociados |

Marcar claramente los datos de demostración. Las integraciones deben ser componentes separados para poder añadir otras fuentes después. La integración con Microsoft Lists queda como opción futura, no como requisito de la primera versión.

## 8. Criterios para comprobar el resultado

- Tarjetas y tabla muestran el mismo conjunto al aplicar filtros.
- Los cambios guardados persisten tras recargar la aplicación.
- Las instituciones se registran desde Administración y aparecen en los selectores y filtros sin modificar el código. Las desactivadas conservan su historial y no se ofrecen para nuevas asociaciones.
- Un QA puede participar en varios proyectos y el cálculo respeta los periodos de solapamiento.
- Gantt y calendario reflejan las fechas registradas.
- Las métricas de Squash coinciden con el alcance seleccionado en la herramienta fuente.
- Un fallo de integración se muestra como error o dato desactualizado, nunca como cero casos.
- Los permisos impiden modificar datos a los lectores.
- Credenciales y tokens permanecen en el servidor y fuera del repositorio.

## 9. Instrucción para comenzar con Codex

> Lee INSTRUCTIVO.md y trabaja dentro de la carpeta portafolio-qa. Usa ASP.NET Core con Razor Pages, Bootstrap, JavaScript, SQL Server y Entity Framework Core, en un único proyecto. Define la estructura mínima y una versión LTS de .NET con soporte vigente. Implementa inicialmente solo la etapa 1 con datos de demostración, sin conexiones reales ni publicación. Las instituciones deben registrarse desde la aplicación: inicia su catálogo vacío y no incluyas nombres predefinidos en código, ejemplos ni datos de demostración. Mantén el código preparado para asignaciones QA e integraciones desde el backend. Documenta cómo ejecutar la aplicación en README.md y verifica filtros, navegación y formularios. No inventes endpoints ni credenciales de Squash TM o Bugzilla.
