# Actividades predeterminadas — entrega y verificación

Implementado y aplicado a la base local `PortafolioQA` de `.\SQLEXPRESS` el 9 de octubre de 2026. La configuración efectiva procede de user-secrets. No se desplegó en IIS.

## Implementación

- Migración `20261009185340_Etapa11ActividadesPredeterminadas`: marca de predeterminada, orden, peso (10 para extras), avance heredado y confirmación pendiente.
- Alta y CSV crean las tres actividades 30/50/20, sin responsable y con fechas proporcionales inclusivas. Para rangos menores de tres días se comparte algún día.
- Avance ponderado sobre actividades no canceladas; Estado y Avance del proyecto solo se muestran. Cancelar y reactivar son acciones separadas.
- El avance de cada actividad deriva su estado; Cancelada se conserva hasta reactivar. Los datos heredados En curso al 0 % también mantienen al proyecto En certificación.
- No se borran, renombran ni desprotegen las predeterminadas desde los handlers o `SaveChanges`. Fechas y responsable siguen editables; el peso es interno (30/50/20 y 10 en extras) y ya no se muestra ni se edita en pantalla. Se mantiene la restricción existente: corregir actividades de proyectos cerrados requiere Administrador.
- La confirmación del PM exige un porcentaje válido para **todas** las actividades en un solo envío. Editar una actividad individual no libera el avance heredado. `AvanceLegado` se conserva como referencia y cada confirmación se audita.
- Si no hay actividades vigentes, el avance resulta 0; los cancelados conservan su avance. El redondeo nunca cierra anticipadamente un proyecto que aún no alcanza el 100 % real.
- El arranque reconcilia avance/estado mediante el servicio de cierre antes de reparar coherencia, para conservar fotos históricas y restaurar asignaciones al reabrir.

## Respaldos

Ambos respaldos son `COPY_ONLY` con `CHECKSUM` y pasaron `RESTORE VERIFYONLY`:

```text
C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup\PortafolioQA_PreEtapa11_20261009.bak
C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup\PortafolioQA_PreEtapa11_Final_20261009.bak
```

El segundo se tomó con la aplicación detenida, inmediatamente antes de migrar. Para deshacer el backfill se debe restaurar ese respaldo y usar el código anterior: `Down` elimina columnas, pero no reconstruye los datos históricos modificados.

## Conteos y cambios efectivos

| Estado | Antes | Con actividades antes | Después |
|---|---:|---:|---:|
| En certificación | 3 | 3 | 5 |
| Cerrado | 4 | 2 | 2 |
| Planificado | 0 | 0 | 0 |
| Cancelado | 0 | 0 | 0 |
| Total | 7 | 5 | 7 |

| Proyecto ID | Estado posterior | Avance posterior | Actividades totales | Predeterminadas |
|---|---|---:|---:|---:|
| 1 | En certificación | 1.67 % | 5 | 3 |
| 2 | En certificación | 4.55 % | 5 | 3 |
| 3 | Cerrado | 100 % | 3 | 3 |
| 4 | En certificación | 9.09 % | 4 | 3 |
| 5 | En certificación | 12.50 % | 5 | 3 |
| 6 | Cerrado | 100 % | 3 | 3 |
| 7 | En certificación | 23.08 % | 6 | 3 |

Los proyectos 4 y 7 se reabrieron porque ya tenían actividades y recibieron las predeterminadas faltantes al 0 %, como establece el plan. Sus cierres pasaron a historial y sus asignaciones se restauraron. Los proyectos 3 y 6, cerrados sin actividades, recibieron las tres completadas y mantuvieron el cierre. No hay proyectos reales pendientes de confirmación: todos los En certificación ya tenían actividades.

La consulta posterior confirmó **0 asignaciones con marcas de cierre en proyectos abiertos**. La aplicación local se reinició en `http://localhost:5011`; `/healthz` responde `ok` y la página de login responde HTTP 200.

## Pruebas en copia

Se restauraron copias separadas: `PortafolioQA_Verificacion11` y `PortafolioQA_Verificacion11_Final`. La segunda validó la versión final del backfill. El ejecutor añade fixtures solo en la copia para cubrir estados ausentes en la base real.

Los **9 escenarios del plan pasaron**, incluyendo alta real, porcentajes 30/55, En curso al 0 %, cierre con foto y recorte, reapertura con restauración, protección de predeterminadas, cancelar/reactivar, importación CSV real y backfill. También se verificaron rangos de fechas de 1–100 días, exclusión de canceladas, peso editable, entradas numéricas inválidas y rechazo de confirmación parcial. La compilación Debug y Release terminó con 0 advertencias y 0 errores.

Los handlers se ejecutaron directamente con un contexto HTTP de prueba y usuario Administrador; las vistas Razor se verificaron por compilación. No se realizó una revisión visual completa del navegador.

Resultado del ejecutor: [Verification/resultado-final.txt](Verification/resultado-final.txt).

Para repetir, restaurar una **nueva** copia del respaldo previo con nombre `PortafolioQA_Verificacion...` y ejecutar:

```powershell
dotnet run --project Verification/Verification.csproj --configuration Release -- 'Server=.\SQLEXPRESS;Database=PortafolioQA_VerificacionNueva;Trusted_Connection=True;TrustServerCertificate=True'
```

El ejecutor rechaza bases sin ese prefijo y copias que ya tengan aplicada la migración. Nunca usa ni modifica la base principal para las pruebas.
