# Requisitos de versión d80a1a1

Revisión: 19/10/2026, America/Lima. No acredita instalación ni esquema aplicado.
Objetivo confirmado en master remoto: d80a1a17b8a43fc502d27d4590f11d0a200f0dab.

## Bases de comparación

Contar versiones anteriores por la línea principal: git log --first-parent master.

| Origen | Situación | Acción |
|---|---|---|
| 54e1f13a245504f0785434dbd3816ce808456e7b | Anterior de master: merge de integración | Actualización; Etapa18 si falta |
| eed68879cebb45b6b788090a0890a1907a2d5cb4 | Dos anteriores de master; Portafolio no existía en ese árbol | Instalación nueva; esquema completo hasta Etapa18 |
| 3bf6f8bfc0f0d254cc677df8563033c46f18bb08 | Commit de la rama integrada, aparece en log normal | No equivale a dos anteriores de master por primer padre |

Comparaciones en GitHub:
[anterior a objetivo](https://github.com/mahurtad/PyPrys/compare/54e1f13...d80a1a1),
[dos anteriores a objetivo](https://github.com/mahurtad/PyPrys/compare/eed6887...d80a1a1).

Desde la raíz del proyecto, para comparar árboles exactos:

```powershell
git diff 54e1f13 d80a1a1 -- .
git diff eed6887 d80a1a1 -- .
```

## Actualización desde 54e1f13

| Componente | Requisito | Responsable / evidencia |
|---|---|---|
| Portafolio | Reemplazar archivos: dashboard, navegación Tabler y módulos Squash | Operador; paquete identificado y log |
| SQL Server de Portafolio | Aplicar 20261017003925_Etapa18SquashModulos si falta; crea SquashModulos y SquashModuloCaso con relaciones e índices | DBA o administrador autorizado; historial antes/después |
| Configuración | El diff no agrega claves nuevas; verificar URL y credenciales del ambiente | Administrador; prueba sin exponer secretos |
| Squash TM | Actualizar desde Squash para cargar suites/módulos y casos | Administrador/editor; comprobar filtros de campaña, iteración y módulo |
| Bugzilla | Verificar acceso a producto y lectura de defectos con credenciales existentes | Administrador; resultado funcional |
| Bases Squash/PostgreSQL y Bugzilla/MariaDB | Este salto no cambia Compose, imágenes ni scripts de sus bases; no requiere migrarlas | Infraestructura; inventario |
| Instalador | No cambia; Actualizar reemplaza archivos, no aplica migraciones | Operador; log |

El dashboard lee datos ya sincronizados; abrirlo no sincroniza Squash.
Validar módulos con un proyecto abierto que tenga suites. El sincronizador bloquea proyectos Cerrados/Cancelados: registrar la limitación sin reabrirlos automáticamente.

## Instalación nueva o integración aún no preparada

Seguir [despliegue IIS](../../DEPLOY_IIS.md), [instalador](../../Instalador/LEEME.txt),
[integración Bugzilla](../../INTEGRACION_BUGZILLA.md) e [infraestructura local](../../infra/local/README.md).

- IIS, Hosting Bundle .NET 8, HTTPS y carpeta persistente de claves para servidor.
- SQL Server, base y permisos; configurar /Setup y aplicar TODAS las migraciones pendientes hasta Etapa18.
- Configurar Squash__BaseUrl, Squash__Token, Bugzilla__BaseUrl y Bugzilla__ApiKey en servidor; claves con dos puntos en user-secrets para desarrollo.
- Para productos automáticos, Etapa17 aplicada y cuenta de servicio Bugzilla con editcomponents y Bugzilla__AdminApiKey independiente; Bugzilla__CuentaServicio si corresponde.
- Configurar bugtracker en Squash y asociar manualmente proyecto con producto Bugzilla; verificar componente General y versión 1.0.
- El instalador de Portafolio no instala Squash ni Bugzilla ni actualiza sus bases.
- Tras primer ingreso IIS, cambiar contraseña inicial y retirar Seguridad__AdminInicial__*.

## Orden de ejecución

1. Confirmar origen real, conexión efectiva e historial SQL; guardar respaldo de base, configuración y claves.
2. Generar paquete mediante Instalador/Construir-Instalador.ps1 y registrar SHA-256. Si hay cambios locales, no declararlo build limpio del objetivo.
3. Instalar/actualizar archivos según corresponda.
4. Aplicar migraciones desde Administración > Conexión a la base de datos, o entregar al DBA el SQL idempotente db/portafolio-qa.sql del ZIP.
5. Verificar que no queden pendientes y /healthz devuelva 200 con ok.
6. Validar login, Dashboard, sincronización/módulos Squash, defectos Bugzilla y asociación desde Squash.
7. Registrar resultado y evidencias en la ficha del ambiente.

Ante fallo, registrar incidencia y coordinar recuperación de archivos, configuración y respaldo SQL.
No ejecutar Down de migraciones automáticamente: Etapa18 elimina sus tablas al revertirse.
