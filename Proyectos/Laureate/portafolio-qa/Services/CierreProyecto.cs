using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Services;

/// <summary>Reglas de cierre y reapertura de proyectos, y avance/estado derivados de las actividades.</summary>
public static class CierreProyecto
{
    /// <summary>
    /// Fija el avance en 100, libera capacidad y deja trazabilidad: guarda una foto del cierre (equipo y actividades),
    /// recorta a la fecha de cierre las asignaciones en curso y marca como no ejecutadas las que no habían empezado.
    /// Las actividades conservan su responsable. Devuelve un resumen para mostrar al usuario.
    /// </summary>
    public static async Task<string> AplicarAsync(PortafolioQAContext context, Proyecto proyecto, DateTime fechaCierre, string usuario = UsuarioSistema)
    {
        proyecto.Avance = 100;
        var cierre = fechaCierre.Date;

        var asignaciones = await context.Asignaciones.Include(a => a.RecursoQA)
            .Where(a => a.ProyectoId == proyecto.Id && !a.NoEjecutada).ToListAsync();
        var actividades = await context.Actividades.Include(a => a.RecursoQA)
            .Where(a => a.ProyectoId == proyecto.Id).ToListAsync();

        // Un cierre vigente previo (no debería existir) pasa a historial.
        foreach (var previo in await context.Cierres.Where(c => c.ProyectoId == proyecto.Id && c.Vigente).ToListAsync())
        {
            previo.Vigente = false;
        }

        var registro = new CierreProyectoRegistro
        {
            ProyectoId = proyecto.Id,
            Fecha = DateTime.Now,
            Usuario = usuario,
            AvanceFinal = 100,
            TotalActividades = actividades.Count,
            ActividadesCompletadas = actividades.Count(a => a.Estado == EstadoActividad.Completada),
        };

        int acortadas = 0, noEjecutadas = 0;
        foreach (var a in asignaciones)
        {
            var ejecutada = a.FechaInicio.Date < cierre;
            registro.Participantes.Add(new CierreParticipante
            {
                RecursoQAId = a.RecursoQAId,
                NombreRecurso = a.RecursoQA?.NombreCompleto ?? string.Empty,
                Rol = a.Rol,
                Porcentaje = a.Porcentaje,
                FechaInicio = a.FechaInicio.Date,
                FechaFin = ejecutada && a.FechaFin.Date >= cierre ? cierre.AddDays(-1) : a.FechaFin.Date,
                Ejecutada = ejecutada
            });

            if (a.FechaFin.Date < cierre) continue;
            if (ejecutada)
            {
                a.FechaFinOriginal ??= a.FechaFin;
                a.FechaFin = cierre.AddDays(-1);
                acortadas++;
            }
            else
            {
                a.NoEjecutada = true;
                noEjecutadas++;
            }
        }

        foreach (var act in actividades)
        {
            registro.Actividades.Add(new CierreActividad
            {
                Nombre = act.Nombre,
                Responsable = act.RecursoQA?.NombreCompleto,
                Estado = act.Estado.Texto(),
                Avance = act.Avance,
                FechaInicio = act.FechaInicio,
                FechaFin = act.FechaFin
            });
        }

        context.Cierres.Add(registro);
        return $" Cierre registrado por {usuario}: {registro.Participantes.Count} participante(s) y {actividades.Count} actividad(es) guardados como historial."
               + (acortadas > 0 ? $" {acortadas} asignación(es) finalizada(s) el {cierre.AddDays(-1):dd/MM/yyyy}." : "")
               + (noEjecutadas > 0 ? $" {noEjecutadas} asignación(es) futura(s) marcadas como no ejecutadas." : "");
    }

    /// <summary>
    /// Al reabrir un proyecto cerrado: el cierre vigente pasa a historial y se restauran las asignaciones
    /// que el cierre había recortado o marcado como no ejecutadas (revisar la capacidad en Planificación).
    /// </summary>
    public static async Task<string> ReabrirAsync(PortafolioQAContext context, Proyecto proyecto, string usuario = UsuarioSistema)
    {
        foreach (var c in await context.Cierres.Where(c => c.ProyectoId == proyecto.Id && c.Vigente).ToListAsync())
        {
            c.Vigente = false;
            c.FechaReapertura = DateTime.Now;
            c.UsuarioReapertura = usuario;
        }

        var restauradas = 0;
        foreach (var a in await context.Asignaciones.Where(a => a.ProyectoId == proyecto.Id && (a.NoEjecutada || a.FechaFinOriginal != null)).ToListAsync())
        {
            a.NoEjecutada = false;
            if (a.FechaFinOriginal is { } original) a.FechaFin = original;
            a.FechaFinOriginal = null;
            restauradas++;
        }
        return restauradas > 0 ? $" Se restauraron {restauradas} asignación(es): revisa la capacidad de los recursos." : string.Empty;
    }

    /// <summary>Registra un cambio sobre una actividad (se guarda con el siguiente SaveChanges).</summary>
    public static void Auditar(PortafolioQAContext context, Proyecto proyecto, Actividad? actividad, string nombre, string usuario, string accion, string? detalle)
    {
        context.ActividadAuditorias.Add(new ActividadAuditoria
        {
            ProyectoId = proyecto.Id,
            ActividadId = actividad?.Id,
            Actividad = nombre,
            Fecha = DateTime.Now,
            Usuario = usuario,
            Accion = accion,
            Detalle = detalle,
            ProyectoCerrado = proyecto.Estado == EstadoProyecto.Cerrado
        });
    }

    public const string UsuarioSistema = "sistema";

    /// <summary>
    /// Invariantes entre proyecto, asignaciones y cierre, que se verifican al iniciar la app:
    /// un proyecto Cerrado tiene un cierre vigente y ninguna asignación activa más allá de su fecha de cierre;
    /// un proyecto abierto no conserva asignaciones recortadas o no ejecutadas por un cierre anterior.
    /// Devuelve la descripción de cada corrección hecha (vacío si todo estaba coherente).
    /// </summary>
    public static async Task<List<string>> ReparaCoherenciaAsync(PortafolioQAContext context)
    {
        var hoy = DateTime.Today;
        var cambios = new List<string>();

        foreach (var p in await context.Proyectos.Where(p => p.Estado == EstadoProyecto.Cerrado).ToListAsync())
        {
            if (p.AvancePendienteConfirmar) continue;
            foreach (var act in await context.Actividades.Where(a => a.ProyectoId == p.Id && a.Estado != EstadoActividad.Cancelada).ToListAsync())
            {
                if (act.Avance == 100 && act.Estado == EstadoActividad.Completada) continue;
                act.Avance = 100;
                act.Estado = EstadoActividad.Completada;
                cambios.Add($"'{p.Nombre}': actividad '{act.Nombre}' alineada con el cierre al 100 %.");
            }
            var vigente = await context.Cierres.Include(c => c.Participantes)
                .Where(c => c.ProyectoId == p.Id && c.Vigente).OrderByDescending(c => c.Fecha).FirstOrDefaultAsync();
            if (vigente is null)
            {
                await AplicarAsync(context, p, hoy, "Corrección de coherencia");
                cambios.Add($"'{p.Nombre}': estaba Cerrado sin registro de cierre; se registró y se liberaron sus recursos.");
                continue;
            }

            // Un cierre no puede estar en el futuro (p. ej. fechas reconstruidas): se toma hoy como fecha efectiva.
            if (vigente.Fecha.Date > hoy)
            {
                vigente.Fecha = hoy;
                cambios.Add($"'{p.Nombre}': la fecha de cierre estaba en el futuro; se fijó en {hoy:dd/MM/yyyy}.");
            }
            var corte = vigente.Fecha.Date;

            // Toda asignación del proyecto debe figurar en la foto del cierre (p. ej. una creada después de la reconstrucción).
            foreach (var a in await context.Asignaciones.Include(a => a.RecursoQA)
                         .Where(a => a.ProyectoId == p.Id && !a.NoEjecutada).ToListAsync())
            {
                if (vigente.Participantes.Any(x => x.RecursoQAId == a.RecursoQAId && x.Rol == a.Rol)) continue;
                var fin = (a.FechaFinOriginal ?? a.FechaFin).Date >= corte ? corte.AddDays(-1) : a.FechaFin.Date;
                vigente.Participantes.Add(new CierreParticipante
                {
                    RecursoQAId = a.RecursoQAId,
                    NombreRecurso = a.RecursoQA?.NombreCompleto ?? string.Empty,
                    Rol = a.Rol,
                    Porcentaje = a.Porcentaje,
                    FechaInicio = a.FechaInicio.Date,
                    FechaFin = a.FechaInicio.Date < corte ? fin : a.FechaFin.Date,
                    Ejecutada = a.FechaInicio.Date < corte
                });
                cambios.Add($"'{p.Nombre}': se agregó al historial de cierre a {a.RecursoQA?.NombreCompleto}, que no figuraba.");
            }

            var pendientes = await context.Asignaciones
                .Where(a => a.ProyectoId == p.Id && !a.NoEjecutada && a.FechaFin >= corte).ToListAsync();
            foreach (var a in pendientes)
            {
                var ejecutada = a.FechaInicio.Date < corte;
                if (ejecutada)
                {
                    a.FechaFinOriginal ??= a.FechaFin;
                    a.FechaFin = corte.AddDays(-1);
                }
                else
                {
                    a.NoEjecutada = true;
                }

                var part = vigente.Participantes.FirstOrDefault(x => x.RecursoQAId == a.RecursoQAId && x.Rol == a.Rol);
                if (part is not null)
                {
                    part.Ejecutada = ejecutada;
                    if (ejecutada) part.FechaFin = a.FechaFin;
                }
                cambios.Add($"'{p.Nombre}': asignación de recurso {a.RecursoQAId} " + (ejecutada
                    ? $"recortada al {a.FechaFin:dd/MM/yyyy} (fecha de cierre {corte:dd/MM/yyyy})."
                    : "marcada como no ejecutada."));
            }
        }

        foreach (var c in await context.Cierres.Where(c => !c.Vigente && c.Fecha > hoy).ToListAsync())
        {
            c.Fecha = hoy;
            cambios.Add($"Cierre histórico del proyecto {c.ProyectoId}: su fecha estaba en el futuro; se fijó en {hoy:dd/MM/yyyy}.");
        }

        var abiertos = await context.Proyectos
            .Where(p => p.Estado != EstadoProyecto.Cerrado && p.Estado != EstadoProyecto.Cancelado).Select(p => p.Id).ToListAsync();
        foreach (var a in await context.Asignaciones
                     .Where(a => abiertos.Contains(a.ProyectoId) && (a.NoEjecutada || a.FechaFinOriginal != null)).ToListAsync())
        {
            a.NoEjecutada = false;
            if (a.FechaFinOriginal is { } original) a.FechaFin = original;
            a.FechaFinOriginal = null;
            cambios.Add($"Asignación {a.Id}: pertenecía a un proyecto abierto pero conservaba marcas de cierre; se restauró.");
        }

        if (cambios.Count > 0) await context.SaveChangesAsync();
        return cambios;
    }

    /// <summary>True si el proyecto no admite cambios de equipo: un proyecto cerrado es historial hasta que se reabra.</summary>
    public static async Task<bool> EquipoBloqueadoAsync(PortafolioQAContext context, int proyectoId)
    {
        var estado = await context.Proyectos.Where(p => p.Id == proyectoId).Select(p => (EstadoProyecto?)p.Estado).FirstOrDefaultAsync();
        return estado is EstadoProyecto.Cerrado;
    }

    /// <summary>Estado que corresponde a un avance: 0 = Planificado, 1-99 = En certificación, 100 = Cerrado.</summary>
    public static EstadoProyecto EstadoPorAvance(decimal avance) =>
        avance >= 100 ? EstadoProyecto.Cerrado
        : avance <= 0 ? EstadoProyecto.Planificado
        : EstadoProyecto.EnCertificacion;

    public static EstadoActividad EstadoActividadPorAvance(decimal avance) =>
        avance >= 100 ? EstadoActividad.Completada
        : avance <= 0 ? EstadoActividad.Pendiente : EstadoActividad.EnCurso;

    /// <summary>Avance ponderado sobre actividades no canceladas. El legado pendiente se conserva.</summary>
    public static async Task<string> RecalcularAvanceAsync(PortafolioQAContext context, Proyecto proyecto, string usuario = UsuarioSistema)
    {
        if (proyecto.AvancePendienteConfirmar || proyecto.Estado == EstadoProyecto.Cancelado) return string.Empty;

        var actividades = await context.Actividades.Where(a => a.ProyectoId == proyecto.Id)
            .Select(a => new { a.Estado, a.Avance, a.Peso }).ToListAsync();
        var vigentes = actividades.Where(a => a.Estado != EstadoActividad.Cancelada).ToList();
        var pesoTotal = vigentes.Sum(a => a.Peso);
        var promedio = pesoTotal == 0 ? 0 : vigentes.Sum(a => a.Avance * a.Peso) / pesoTotal;
        // Evita cerrar por un redondeo a 100 cuando todavía hay avance pendiente.
        proyecto.Avance = promedio < 100 ? Math.Min(99.99m, Math.Round(promedio, 2)) : 100;
        var aviso = string.Empty;
        var objetivo = EstadoPorAvance(proyecto.Avance);
        if (objetivo == EstadoProyecto.Planificado && vigentes.Any(a => a.Estado == EstadoActividad.EnCurso))
            objetivo = EstadoProyecto.EnCertificacion;
        if (objetivo != proyecto.Estado)
        {
            var previo = proyecto.Estado;
            proyecto.Estado = objetivo;
            if (objetivo == EstadoProyecto.Cerrado)
            {
                proyecto.Avance = 100;
                aviso = " El proyecto llegó al 100 % y pasó a estado Cerrado." + await AplicarAsync(context, proyecto, DateTime.Today, usuario);
            }
            else if (previo == EstadoProyecto.Cerrado)
            {
                aviso = $" Hay avance pendiente ({proyecto.Avance:0.##} %): el proyecto se reabrió en estado {objetivo.Texto()}."
                        + await ReabrirAsync(context, proyecto, usuario);
            }
        }
        await context.SaveChangesAsync();
        return aviso;
    }

    /// <summary>
    /// Tras quitar o cambiar una asignación: si el recurso ya no está asignado al proyecto, sus actividades quedan
    /// sin responsable. Devuelve cuántas actividades se desvincularon.
    /// </summary>
    public static async Task<int> DesvincularActividadesHuerfanasAsync(PortafolioQAContext context, int proyectoId, int recursoId)
    {
        // En un proyecto cerrado las actividades conservan su responsable como historial.
        if (await context.Proyectos.AnyAsync(p => p.Id == proyectoId && p.Estado == EstadoProyecto.Cerrado)) return 0;
        if (await context.Asignaciones.AnyAsync(a => a.ProyectoId == proyectoId && a.RecursoQAId == recursoId && !a.NoEjecutada)) return 0;

        var actividades = await context.Actividades.Where(a => a.ProyectoId == proyectoId && a.RecursoQAId == recursoId).ToListAsync();
        foreach (var a in actividades) a.RecursoQAId = null;
        if (actividades.Count > 0) await context.SaveChangesAsync();
        return actividades.Count;
    }

    /// <summary>Si la actividad termina después que el proyecto, extiende el fin del proyecto. True si lo modificó.</summary>
    public static bool AjustarFinProyecto(Proyecto proyecto, DateTime finActividad)
    {
        if (finActividad.Date <= proyecto.FechaFin.Date) return false;
        proyecto.FechaFin = finActividad.Date;
        return true;
    }
}
