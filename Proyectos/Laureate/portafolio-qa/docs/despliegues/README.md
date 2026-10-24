# Control de versiones y ambientes

Punto central para saber qué requiere una versión y qué se aplicó en cada ambiente.
Una ficha documentada no acredita que el cambio se haya ejecutado.

| Objetivo | Ambiente | Estado | Ficha |
|---|---|---|---|
| d80a1a1 | Servidor IIS en uso (identificación pendiente) | En mantenimiento: usuario reporta 2 migraciones pendientes | [Servidor](d80a1a1-servidor.md) |
| d80a1a1 | Local de desarrollo | Pendiente de inventario y validación | [Local](d80a1a1-local.md) |

Requisitos compartidos: [versión d80a1a1](version-d80a1a1.md).
Para otra versión o ambiente, copiar [PLANTILLA.md](PLANTILLA.md) y agregar una fila.
Cada servidor tiene su propia ficha.

## Procedimiento

1. Confirmar commit instalado, commit objetivo y migraciones aplicadas en la base efectiva.
2. Comparar commits y revisar cambios de esquema, configuración, instalador e integraciones.
3. Registrar respaldo de base, configuración y claves de protección, con ubicación y responsable.
4. Identificar paquete por nombre, commit y SHA-256. El instalador declara 2.0.0: ese número no distingue builds.
5. Registrar responsable, fecha con zona horaria, resultado y evidencia de cada paso.
6. Cerrar como Validado solo después de verificar health, login, dashboard e integraciones.

No incluir contraseñas, tokens ni cadenas de conexión. Referenciar su almacén autorizado.
Los cambios locales sin commit no forman parte del commit objetivo.

## Evidencias

- Git: SHA y comparación de código.
- Paquete: nombre, SHA-256 y origen limpio o detalle de cambios locales.
- SQL Server: resultado antes y después de la consulta siguiente, indicando servidor y base sin credenciales.

```sql
SELECT MigrationId, ProductVersion
FROM dbo.__EFMigrationsHistory
ORDER BY MigrationId;
```

- IIS: archivar el log de la ejecución desde C:\ProgramData\PortafolioQA\instalacion.log.
- Cambios de conexión: C:\ProgramData\PortafolioQA\auditoria-configuracion.log.
- Pruebas: /healthz devuelve 200 y ok; login, Dashboard, sincronización Squash y defectos Bugzilla.
- Local: infra/local/Verificar.ps1 requiere PowerShell 7; complementar con pruebas funcionales.

SQL manual corregido y probado: [PortafolioQA-Servidor-Corregido.sql](PortafolioQA-Servidor-Corregido.sql).
Consultar la incidencia y validación en la ficha del servidor antes de ejecutarlo.
