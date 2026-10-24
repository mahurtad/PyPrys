# Integración Bugzilla: producto automático por iniciativa

Al guardar una iniciativa vinculada a un proyecto de Squash TM, Portafolio guarda primero la iniciativa y luego prepara en
Bugzilla su producto (nombre del proyecto de Squash), el componente `General` y la versión `1.0`. La relación es 1 a 1:
iniciativa ↔ proyecto de Squash ↔ producto de Bugzilla. Todo se hace por la **API REST oficial** de Bugzilla; no se usa SQL,
Perl ni Docker contra Bugzilla desde la aplicación.

Estados que muestra Portafolio: **Sincronizado con Bugzilla** (producto, componente y versión verificados y vínculo guardado),
**Pendiente** y **Error** (con botón *Reintentar* en el detalle de la iniciativa). Un fallo de Bugzilla nunca revierte la iniciativa.

## Qué debe hacer cada rol antes de activarlo

### DBA (base de Portafolio, SQL Server)
- Aplicar la migración `Etapa17VinculoBugzilla`: crea la tabla `VinculosBugzilla` (solo agrega, no modifica ni borra datos).
  Opciones: el asistente `/Setup`, o el script idempotente (`dotnet ef migrations script --idempotent`).
- Si la cuenta de la aplicación no tiene `db_ddladmin`, el DBA ejecuta el script de esquema. Con migraciones desde el asistente
  hace falta `db_ddladmin`; para operar basta `db_datareader` y `db_datawriter`.
- Las iniciativas ya vinculadas a Squash **no** reciben vínculo de Bugzilla automáticamente: se crea al guardarlas o al pulsar
  *Reintentar*. Nada se borra ni se reasigna solo.

### Administrador de Bugzilla
- Crear una **cuenta de servicio dedicada** (p. ej. `portafolio-productos-api@<dominio>`) con **solo el grupo `editcomponents`**.
  Es el permiso que Bugzilla exige para crear productos y componentes (`POST /rest/product`, `POST /rest/component`).
  No usar la cuenta de un administrador humano ni darle más grupos.
- Generar su **API key** (Bugzilla → Preferencias → API Keys, iniciando sesión como esa cuenta). Bugzilla 5.2 no permite crear
  API keys por REST. En la instalación local, `infra/local/Crear-ApiProductos.ps1` lo automatiza (solo local con Docker).
- Esta cuenta es distinta de `Bugzilla:ApiKey` (consulta de defectos): no reutilizar esa clave.

### Quien despliega Portafolio
- **Forma recomendada:** un Administrador de Portafolio configura todo en **Sistema → Integración Bugzilla** (guía paso a paso en
  la pantalla): la **dirección** de Bugzilla, la **clave de lectura** (defectos) y la **clave de la cuenta de servicio** con su login
  (crea productos). Cada dato se prueba contra Bugzilla (la dirección responde como Bugzilla; la clave de servicio es válida, es de
  esa cuenta y tiene `editcomponents`), se guarda cifrado con Data Protection en `bugzilla-servicio.dat` (carpeta de datos del
  servidor, junto a `conexion.dat`) y surte efecto sin reiniciar. Lo guardado tiene prioridad sobre la configuración del servidor;
  *Quitar lo guardado* devuelve a esta. Los cambios se auditan (sin las claves). Guardar exige la contraseña del administrador.
- **Cambiar la dirección** obliga a pegar de nuevo las dos claves (son de otra instalación). Si hay proyectos con producto ya creado
  en la dirección anterior, la pantalla pide confirmación: esos vínculos quedan en Error para revisión (ver reglas de seguridad);
  no se borra ni se reasigna nada.
- **Alternativa por servidor:** lo siguiente sigue funcionando y es lo que rige si no se guardó nada desde la pantalla.
- Configurar la clave **fuera del repositorio**: user-secrets (`Bugzilla:AdminApiKey`) en desarrollo, o la variable de entorno
  `Bugzilla__AdminApiKey` en el servidor. Opcionalmente `Bugzilla:CuentaServicio` (login de esa cuenta; es el responsable por defecto
  del componente `General`; no se asigna a personas).
- `Bugzilla:BaseUrl` y `Squash:BaseUrl` deben ser las de **ese** ambiente: los vínculos guardan los IDs junto con esas URL.
- Sin `Bugzilla:AdminApiKey` el estado queda **Pendiente** (no falla): configurar la clave y pulsar *Reintentar*.

## Particularidades y problemas encontrados
1. **Bugzilla 5.2 ignora la cabecera `X-BUGZILLA-API-KEY`.** Con la cabecera, las lecturas funcionan como anónimas y crear
   devuelve 401 "You must log in". La clave se envía como parámetro `Bugzilla_api_key` (`BugzillaApiKeyHandler`). El log HTTP de ese
   cliente está en nivel Warning para que la URL con la clave no se registre. No subir ese nivel.
   *Corregido (19/10/2026):* el cliente de **lectura** de defectos (`BugzillaClient`) también enviaba la cabecera y consultaba como
   anónimo (comprobado con `valid_login`: cabecera = `false`, parámetro = `true`). Ahora usa el mismo mecanismo y su log HTTP también
   está en Warning.
2. **Una API key no se puede crear por REST.** Hay que generarla desde la interfaz de Bugzilla (o, en local, con el script).
3. **La versión `1.0` solo se crea junto con el producto.** Bugzilla 5.2 no tiene endpoint para agregar versiones a un producto
   existente. Si a un producto ya existente le falta `1.0`, queda en Error con la instrucción de agregarla a mano y reintentar.
4. **El componente exige un responsable por defecto** (`default_assignee`): se usa la cuenta de servicio.
5. **Asociar el proyecto de Squash con el producto de Bugzilla es un paso MANUAL.** Squash TM 11.0.9 no ofrece API para asociar un
   proyecto a un bugtracker ni para leer esa asociación (solo se copia al crear un proyecto desde otro). Hacerlo en Squash →
   Administración → Proyectos → *Bugtracker*. La etiqueta de Portafolio **no** afirma ese paso.

## Reglas de seguridad de los vínculos
- Los IDs se guardan con su ambiente (URL base normalizada): un vínculo creado en otro Squash/Bugzilla no se reutiliza ni se
  reasigna; queda en Error para revisión.
- Un producto existente con el mismo nombre solo se asocia si su descripción lleva la marca de esta iniciativa
  (`[portafolio-qa squash=<url>#<id>]`) o si un administrador lo confirma con *Asociar existente*. Un producto ya vinculado a otra
  iniciativa nunca se reutiliza.
- Una vez vinculado, el producto se identifica por ID: renombrar la iniciativa no cambia el producto. Si el producto desaparece
  de Bugzilla, queda en Error; no se crea otro automáticamente.
- Reintentos y solicitudes simultáneas no duplican: hay reserva atómica en la base e índices únicos.

## Verificación
- Pruebas automatizadas: `dotnet test Tests/PortafolioQA.Tests` (usa bases temporales en `.\SQLEXPRESS`; ajustar la cadena en
  `Entorno` si el SQL Server es otro).
- Prueba de extremo a extremo local (16/10/2026): proyecto Squash "Proyecto Prueba 3" (ID 3) → iniciativa 13 → producto Bugzilla 3,
  componente `General` (3), versión `1.0` (3); guardados repetidos sin duplicados; el proyecto vinculado deja de aparecer en el selector.
- Instalación local de apoyo: `infra/local/README.md`.
