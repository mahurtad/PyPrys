# Plan de implementación — Actividades predeterminadas y avance/estado derivados

## Objetivo

**Estado al 09/10/2026:** cinco fases implementadas; migración aplicada a la BD local tras conteo y respaldo verificado. Los nueve escenarios pasaron sobre una copia. Ver [resultados, respaldos y cambios de estado](VERIFICACION_ETAPA11.md). Sin despliegue en IIS.

Todo proyecto nace con 3 actividades predeterminadas. Su **avance y estado dejan de editarse a mano**: salen de las actividades. Cancelar un proyecto sigue siendo la única acción manual sobre el estado.

## Decisiones tomadas (defaults)

| Tema | Decisión |
|---|---|
| Actividades predeterminadas | 1. Diseño y redacción de Casos de Prueba · 2. Ejecución de Casos de prueba · 3. Elaborar informe |
| Peso | Interno, sin edición ni visibilidad en pantalla: 30 / 50 / 20 en las predeterminadas y 10 en las extras. El usuario solo registra el Avance de cada actividad. Avance = Σ(avance × peso) / Σ(peso) sobre actividades no canceladas |
| Fechas | Secuenciales y proporcionales al rango del proyecto (30 % / 50 % / 20 % de los días calendario), editables |
| Estado del proyecto | 100 % → Cerrado · avance > 0 **o** alguna actividad En curso → En certificación · resto → Planificado · Cancelado manual |
| Estado de la actividad | Derivado del avance: 0 → Pendiente, 1-99 → En curso, 100 → Completada. Cancelada es manual |
| Protección | Las 3 predeterminadas no se borran ni se renombran; sí se editan fechas, responsable y avance |
| Responsable | Las predeterminadas nacen sin responsable; se asigna después entre los QA del proyecto |
| Backfill "En certificación" | No se inventa avance: se marcan para confirmación del PM (ver Fase 4) |

## Fase 1 — Modelo y migración `Etapa11ActividadesPredeterminadas`

- [Models/Actividad.cs](Models/Actividad.cs): agregar `EsPredeterminada` (bool), `Orden` (int), `Peso` (decimal, default 10).
- [Models/Proyecto.cs](Models/Proyecto.cs): agregar `AvancePendienteConfirmar` (bool) y `AvanceLegado` (decimal?) para el backfill.
- Constantes en un nuevo `Services/ActividadesPredeterminadas.cs`: nombres, pesos, porcentajes de fechas y un método `Crear(Proyecto)` que devuelve las 3 con fechas calculadas.
- Migración con `dotnet ef migrations add Etapa11ActividadesPredeterminadas`. El backfill de datos va en SQL dentro de la migración (Fase 4).

## Fase 2 — Reglas de negocio ([Services/CierreProyecto.cs](Services/CierreProyecto.cs))

- `RecalcularAvanceAsync`:
  - avance ponderado por `Peso`; ya no existe el modo manual;
  - estado según la regla nueva (incluye "alguna actividad En curso");
  - si `AvancePendienteConfirmar` es true, no recalcula ni cambia el estado.
- Nuevo `EstadoActividadPorAvance(avance)` y su uso en cada alta o edición de actividad.
- Validar que no se pueda eliminar o renombrar una actividad con `EsPredeterminada`.
- Revisar `ReparaCoherenciaAsync`: un proyecto Cerrado debe tener sus actividades al 100 %.
- Revisar `RecursoValidoAsync` ([Edit.cshtml.cs](Pages/Admin/Proyectos/Edit.cshtml.cs)): permitir `null` en las predeterminadas aunque haya varios QA asignados.

## Fase 3 — Pantallas

- [Pages/Admin/Proyectos/Create.cshtml.cs](Pages/Admin/Proyectos/Create.cshtml.cs): quitar Estado y Avance del binding; crear el proyecto Planificado al 0 % con las 3 actividades. Las fechas se calculan al guardar a partir de Inicio y Fin.
- [Pages/Admin/Proyectos/_Form.cshtml](Pages/Admin/Proyectos/_Form.cshtml): Estado y Avance pasan a texto de solo lectura (badge y barra), con la nota "Se calcula a partir de las actividades". Se elimina el script de sincronización. En edición, un botón separado "Cancelar proyecto" / "Reactivar".
- [Pages/Admin/Proyectos/Edit.cshtml.cs](Pages/Admin/Proyectos/Edit.cshtml.cs): quitar las ramas de avance y estado manual.
  - Eliminar el handler de cambio de estado de actividad. El estado se deriva, salvo la acción Cancelar actividad.
  - Mostrar un candado en las predeterminadas; el peso no se muestra ni se edita.
  - Corregir el texto "terminadas / total", que no coincide con el cálculo.
- [Pages/Admin/Proyectos/Importar.cshtml.cs](Pages/Admin/Proyectos/Importar.cshtml.cs): crear las 3 actividades por cada proyecto importado y actualizar el texto de ayuda.
- [Pages/Portafolio/Detalle.cshtml](Pages/Portafolio/Detalle.cshtml): un aviso si el avance está pendiente de confirmación.
- [Pages/Admin/Proyectos/Index.cshtml](Pages/Admin/Proyectos/Index.cshtml): badge "Avance por confirmar" cuando aplique.

## Fase 4 — Backfill de proyectos existentes

Antes de migrar, ejecutar un conteo en la BD y revisar el resultado:

```sql
SELECT Estado, COUNT(*) AS Proyectos, SUM(CASE WHEN EXISTS (SELECT 1 FROM Actividades a WHERE a.ProyectoId = p.Id) THEN 1 ELSE 0 END) AS ConActividades
FROM Proyectos p GROUP BY Estado;
```

Reglas, solo para proyectos **sin** actividades:

| Estado actual | Qué se crea |
|---|---|
| Cerrado | 3 predeterminadas Completadas al 100 %; no se toca el cierre existente |
| Planificado (0 %) | 3 predeterminadas Pendientes al 0 % |
| En certificación (X %) | 3 predeterminadas al 0 %, `AvancePendienteConfirmar = 1`, `AvanceLegado = X`. El avance mostrado sigue siendo X hasta que el PM confirme |
| Cancelado | 3 predeterminadas Canceladas; el estado no se modifica |

Proyectos que ya tienen actividades: se les agrega `EsPredeterminada`/`Orden` solo si el nombre coincide; si no, se crean las 3 que falten y el avance se recalcula.

**Confirmación del PM:** pantalla en Edit con el aviso "Avance heredado: X %". El PM ingresa el avance de cada actividad y se guarda, lo que limpia `AvancePendienteConfirmar`. Hasta entonces el proyecto no cambia de estado ni se cierra.

## Fase 5 — Verificación

Contra una BD de prueba (copia, no producción):

1. Crear proyecto → 3 actividades, 30/50/20, fechas secuenciales dentro del rango, Planificado 0 %.
2. Diseño al 100 % → avance 30 %, En certificación. Ejecución al 50 % → 55 %.
3. Actividad En curso con 0 % → proyecto pasa a En certificación.
4. Las 3 al 100 % → Cerrado, con foto de cierre y asignaciones recortadas.
5. Agregar una actividad extra al proyecto cerrado → se reabre y se restauran las asignaciones.
6. Intentar borrar o renombrar una predeterminada → rechazado.
7. Cancelar y reactivar el proyecto → conserva el avance.
8. Importación CSV → cada proyecto con sus 3 actividades.
9. Backfill sobre una copia de producción → cuadrar conteos por estado antes y después.

## Orden de entrega y riesgo

1. Fases 1-2 con la migración, sin tocar pantallas (reversible).
2. Fase 3.
3. Fase 4 en una ventana controlada, con respaldo previo de la BD. Es el único paso difícil de revertir.
4. Fase 5 antes de desplegar en IIS (ver [DEPLOY_IIS.md](DEPLOY_IIS.md)).

## Fuera de alcance

- Alimentar "Ejecución de casos" automáticamente desde Squash TM (la integración sigue pendiente).
- Mostrar actividades en el Gantt.
