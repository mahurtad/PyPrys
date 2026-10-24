# Instalación local de QA

Servicios exclusivos del proyecto Compose `portafolio-qa-local`:

- Squash TM 11.0.9: http://localhost:8090/squash (PostgreSQL 17).
- Bugzilla 5.2, rama oficial release-5.2-stable: http://localhost:8080 (MariaDB 11.8).
- Portafolio: http://localhost:5011.

## Inicio y persistencia

Para iniciar las tres aplicaciones con un solo comando, abrir Docker Desktop y ejecutar
`./Iniciar-Todo.ps1` desde la raíz del proyecto. Portafolio queda ejecutándose en esa ventana.

Iniciar Docker Desktop y ejecutar `./infra/local/Iniciar.ps1` desde el proyecto.
Ejecutar `./reiniciar.ps1` para iniciar o recompilar Portafolio.
Los datos persisten en volúmenes Docker; `docker compose stop` conserva todos los datos.
No ejecutar `down -v`: elimina las bases y los adjuntos.
Los puertos web solo se publican en 127.0.0.1; las bases no publican puertos.

Si Docker no puede montar `start-plugins.cfg` desde la unidad virtual de Google Drive,
copiar ese archivo a una carpeta local del equipo y agregar a `.env`
`SQUASH_PLUGINS_CONFIG=C:/ruta/local/start-plugins.cfg`.
Compose usa esa ruta cuando está definida; en los demás equipos conserva la ruta relativa.

## Configuración

`appsettings.json` apunta a estas instalaciones locales. Las credenciales de API
se guardan en `dotnet user-secrets` del proyecto, nunca en el repositorio:

- `Squash:BaseUrl`, `Squash:Token`
- `Bugzilla:BaseUrl`, `Bugzilla:ApiKey`

Las claves de las bases y el administrador inicial Bugzilla viven en `.env`, ignorado por Git.
La conexión SQL Server de Portafolio es independiente y se conserva.

## Asociar iniciativas

En Administración → Proyectos, seleccionar el proyecto de la nueva instalación de
Squash TM y usar Actualizar desde Squash para consultar campañas, iteraciones y casos.
En Administración → Integraciones, seleccionar Bugzilla e indicar el **nombre exacto
del producto**. El detalle de la iniciativa consulta sus defectos y muestra estado,
severidad, responsable y enlaces. Los fallos se muestran como no disponibles.

Las instalaciones nuevas no contienen los proyectos del servidor remoto. No se
migran datos ni se reasignan vínculos automáticamente: los IDs de servidores
independientes pueden coincidir sin representar el mismo proyecto.

## Referencias

- https://tm-en.doc.squashtest.com/v11/install-guide/installation/install-squash/docker-install.html
- https://www.bugzilla.org/download/
- https://bugzilla.readthedocs.io/en/latest/api/core/v1/bug.html

La imagen Bugzilla reutiliza los archivos de referencia locales del proyecto
Personales/Bugzilla/bugzilla-docker/app y descarga el código oficial durante el build.

## Verificación y accesos

Ejecutar ./infra/local/Verificar.ps1 para comprobar las API autenticadas y /healthz.
Los accesos web están en credentials.local.txt (excluido de Git).
El token Squash de lectura vence el 16/10/2027 y se renueva en la cuenta de Squash.


## Conexión directa Squash TM → Bugzilla

Activada en Squash: Administración → Bugtrackers, servidor **Bugzilla local Portafolio QA**.
URL del conector: http://bugzilla.localhost:8080. Los navegadores resuelven `.localhost`
al equipo local; Docker usa el alias de red configurado en Compose para llegar a Bugzilla.
Las API de Portafolio conservan http://localhost:8080.

Squash usa autenticación básica compartida con la cuenta de servicio
`squash-connector@portafolio.local` (editbugs, canconfirm). Los accesos están en
credentials.local.txt, excluido de Git.

El proyecto Squash **Proyecto Prueba 1 (ID 1)** está asociado al producto Bugzilla
**Proyecto Prueba 1**, componente **General**, versión **1.0**. Para otros proyectos,
asociar el mismo servidor y seleccionar su producto desde la administración de Squash.
Se verificó el estado AUTHENTICATED desde el panel de incidencias de Squash.

## Producto de Bugzilla por iniciativa (automático)

Al guardar una iniciativa vinculada a Squash TM, Portafolio guarda primero la iniciativa y luego prepara su producto
en Bugzilla **solo por la API REST oficial**: producto con el nombre del proyecto de Squash, componente `General` y
versión inicial `1.0`. Es idempotente, reanudable y se puede reintentar desde Portafolio → Detalle (Reintentar).

- Credencial: cuenta de servicio `portafolio-productos-api@portafolio.local` (grupo `editcomponents`, nada más).
  Crearla una vez con `./infra/local/Crear-ApiProductos.ps1`; la clave queda en user-secrets
  (`Bugzilla:AdminApiKey`) y en `credentials.local.txt`. En servidor: variable `Bugzilla__AdminApiKey`.
  `Bugzilla:ApiKey` (solo consulta de defectos) no cambia. Sin `AdminApiKey` el vínculo queda **Pendiente**.
- El vínculo guarda los IDs de Squash y Bugzilla con su ambiente (URL base). Un vínculo creado en otro ambiente no se
  reutiliza ni se reasigna. El producto se identifica por ID: renombrar la iniciativa no lo cambia.
- Un producto preexistente con el mismo nombre solo se asocia si su descripción lleva la marca de esta iniciativa o si un
  administrador lo confirma con «Asociar existente».
- Bugzilla 5.2 solo crea la versión junto con el producto y no tiene API para agregar versiones a uno existente.
- **Paso manual:** Squash TM 11.0.9 no ofrece API para asociar un proyecto a un bugtracker/producto (solo se copia al
  crear un proyecto desde otro). Hay que hacerlo en Squash → Administración → Proyectos → Bugtracker. La etiqueta
  «Sincronizado con Bugzilla» de Portafolio no afirma ese paso.
- Pruebas: `dotnet test Tests/PortafolioQA.Tests` (usa bases temporales en `.\SQLEXPRESS`).
