# Ficha d80a1a1 / local de desarrollo

| Dato | Valor |
|---|---|
| Estado | Planificado; servicios y base no verificados en esta revisión |
| Ambiente / equipo / responsable | Local de desarrollo; identificación pendiente |
| Objetivo | d80a1a17b8a43fc502d27d4590f11d0a200f0dab |
| Origen instalado / base / historial SQL | Por confirmar |
| Requisitos | [Versión y escenarios](version-d80a1a1.md) |
| URL previstas | Portafolio http://localhost:5011; Squash http://localhost:8090/squash; Bugzilla http://localhost:8080 |
| Compose previsto | Squash 11.0.9/PostgreSQL 17; Bugzilla rama 5.2/MariaDB 11.8; versiones ejecutadas por confirmar |
| Artefacto | Fuentes locales; hay cambios sin commit, no se certifica build limpio del objetivo |
| Fecha de preparación | 19/10/2026, America/Lima; ejecución pendiente |
| Respaldo / responsable / recuperación | Pendiente |

| Paso | Resultado / evidencia |
|---|---|
| Confirmar origen, base, migraciones y versiones ejecutadas | Pendiente |
| Respaldar base y configuración | Pendiente |
| Iniciar Docker y aplicaciones | Pendiente; Iniciar-Todo.ps1 disponible como cambio local fuera del objetivo |
| Aplicar pendientes hasta Etapa18 | Pendiente |
| Comprobar user-secrets del ambiente sin copiar valores | Pendiente |
| Sincronizar Squash y verificar módulos | Pendiente |
| Verificar Bugzilla y asociación desde Squash | Pendiente |
| Health, login y Dashboard | Pendiente; complementar pwsh -File ./infra/local/Verificar.ps1 con pruebas funcionales |

No se ejecutaron servicios ni migraciones al preparar esta ficha.
