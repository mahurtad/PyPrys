using System.Globalization;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Planificacion;

public class CalBarra
{
    public int ColInicio { get; set; }   // 1..7
    public int ColFin { get; set; }      // 1..7 (inclusive)
    public int Carril { get; set; }
    public string Texto { get; set; } = "";
    public string Titulo { get; set; } = "";
    public string Clase { get; set; } = "";
    public bool ContinuaIzquierda { get; set; }
    public bool ContinuaDerecha { get; set; }
}

public class CalDia
{
    public DateTime Fecha { get; set; }
    public bool EnMes { get; set; }
    public bool EsHoy { get; set; }
    public bool NoLaborable { get; set; }
    public string? Feriado { get; set; }
    public List<string> Hitos { get; set; } = new();
}

public class CalSemana
{
    public List<CalDia> Dias { get; set; } = new();
    public List<CalBarra> Barras { get; set; } = new();
    public int Carriles => Barras.Count == 0 ? 0 : Barras.Max(b => b.Carril) + 1;
}

public class CalendarioModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public CalendarioModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public DateTime Mes { get; set; }
    public int? RecursoId { get; set; }
    public int? ProyectoId { get; set; }
    public List<RecursoQA> Recursos { get; set; } = new();
    public List<Proyecto> Proyectos { get; set; } = new();
    public List<CalSemana> Semanas { get; set; } = new();
    public string TituloMes => Mes.ToString("MMMM yyyy", CultureInfo.GetCultureInfo("es-PE"));
    public string MesAnterior => Mes.AddMonths(-1).ToString("yyyy-MM");
    public string MesSiguiente => Mes.AddMonths(1).ToString("yyyy-MM");
    public int TotalEventos { get; set; }

    public async Task OnGetAsync(string? mes, int? recursoId, int? proyectoId)
    {
        Mes = DateTime.TryParseExact(mes, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var m)
            ? new DateTime(m.Year, m.Month, 1)
            : new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        RecursoId = recursoId;
        ProyectoId = proyectoId;

        Recursos = await _context.RecursosQA.AsNoTracking().Where(r => r.Activo).OrderBy(r => r.Nombres).ToListAsync();
        Proyectos = await _context.Proyectos.AsNoTracking().Where(p => p.Activo).OrderBy(p => p.Nombre).ToListAsync();

        // Rejilla de lunes a domingo que cubre el mes completo.
        var primerDia = Mes;
        var offset = ((int)primerDia.DayOfWeek + 6) % 7;
        var inicioGrid = primerDia.AddDays(-offset);
        var finMes = Mes.AddMonths(1).AddDays(-1);
        var finOffset = 6 - (((int)finMes.DayOfWeek + 6) % 7);
        var finGrid = finMes.AddDays(finOffset);

        var asignaciones = _context.Asignaciones.AsNoTracking()
            .Include(a => a.Proyecto).Include(a => a.RecursoQA)
            .Where(a => a.Proyecto!.Estado != EstadoProyecto.Cancelado && !a.NoEjecutada
                        && a.FechaFin >= inicioGrid && a.FechaInicio <= finGrid);
        if (RecursoId.HasValue) asignaciones = asignaciones.Where(a => a.RecursoQAId == RecursoId);
        if (ProyectoId.HasValue) asignaciones = asignaciones.Where(a => a.ProyectoId == ProyectoId);
        var listaAsignaciones = await asignaciones.OrderBy(a => a.FechaInicio).ThenBy(a => a.Id).ToListAsync();

        var ausenciasQ = _context.Ausencias.AsNoTracking().Include(a => a.RecursoQA)
            .Where(a => a.FechaFin >= inicioGrid && a.FechaInicio <= finGrid);
        if (RecursoId.HasValue) ausenciasQ = ausenciasQ.Where(a => a.RecursoQAId == RecursoId);
        // Si se filtra por proyecto solo importan las ausencias de quienes están asignados a él.
        if (ProyectoId.HasValue)
        {
            var ids = listaAsignaciones.Select(a => a.RecursoQAId).Distinct().ToList();
            ausenciasQ = ausenciasQ.Where(a => ids.Contains(a.RecursoQAId));
        }
        var ausencias = await ausenciasQ.OrderBy(a => a.FechaInicio).ToListAsync();

        var hitosQ = _context.Hitos.AsNoTracking().Include(h => h.Proyecto)
            .Where(h => h.Fecha >= inicioGrid && h.Fecha <= finGrid && h.Proyecto!.Estado != EstadoProyecto.Cancelado);
        if (ProyectoId.HasValue) hitosQ = hitosQ.Where(h => h.ProyectoId == ProyectoId);
        // Si se filtra por recurso, mostrar solo hitos de proyectos donde participa.
        if (RecursoId.HasValue)
            hitosQ = hitosQ.Where(h => h.Proyecto!.Asignaciones.Any(a => a.RecursoQAId == RecursoId));
        var hitos = await hitosQ.ToListAsync();

        var feriados = await _context.Feriados.AsNoTracking()
            .Where(f => f.Fecha >= inicioGrid && f.Fecha <= finGrid).ToDictionaryAsync(f => f.Fecha.Date, f => f.Nombre);

        TotalEventos = listaAsignaciones.Count + ausencias.Count + hitos.Count;

        var eventos = listaAsignaciones.Select(a => (
                Inicio: a.FechaInicio.Date, Fin: a.FechaFin.Date,
                Texto: $"{a.RecursoQA?.NombreCompleto} · {a.Proyecto?.Nombre} ({a.Porcentaje:0.#}%)",
                Titulo: $"{a.RecursoQA?.NombreCompleto} en {a.Proyecto?.Nombre} ({a.Rol}): {a.FechaInicio:dd/MM/yyyy} – {a.FechaFin:dd/MM/yyyy}, dedicación {a.Porcentaje:0.#}%",
                Clase: "cal-asignacion"))
            .Concat(ausencias.Select(a => (
                Inicio: a.FechaInicio.Date, Fin: a.FechaFin.Date,
                Texto: $"{a.Tipo.Texto()}: {a.RecursoQA?.NombreCompleto}" + (string.IsNullOrWhiteSpace(a.Motivo) ? "" : $" ({a.Motivo})"),
                Titulo: $"{a.Tipo.Texto()} de {a.RecursoQA?.NombreCompleto}: {a.FechaInicio:dd/MM/yyyy} – {a.FechaFin:dd/MM/yyyy}",
                Clase: "cal-ausencia")))
            .ToList();

        for (var inicioSemana = inicioGrid; inicioSemana <= finGrid; inicioSemana = inicioSemana.AddDays(7))
        {
            var finSemana = inicioSemana.AddDays(6);
            var semana = new CalSemana();
            for (var i = 0; i < 7; i++)
            {
                var f = inicioSemana.AddDays(i);
                semana.Dias.Add(new CalDia
                {
                    Fecha = f,
                    EnMes = f.Month == Mes.Month,
                    EsHoy = f == DateTime.Today,
                    NoLaborable = f.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday || feriados.ContainsKey(f),
                    Feriado = feriados.GetValueOrDefault(f),
                    Hitos = hitos.Where(h => h.Fecha.Date == f).Select(h => $"{h.Nombre} · {h.Proyecto?.Nombre}").ToList()
                });
            }

            // Reparto en carriles: cada barra ocupa el primer carril cuyo último extremo no la solape.
            var finPorCarril = new List<int>();
            foreach (var e in eventos.Where(e => e.Fin >= inicioSemana && e.Inicio <= finSemana))
            {
                var colIni = ((e.Inicio > inicioSemana ? e.Inicio : inicioSemana) - inicioSemana).Days + 1;
                var colFin = ((e.Fin < finSemana ? e.Fin : finSemana) - inicioSemana).Days + 1;
                var carril = finPorCarril.FindIndex(fin => fin < colIni);
                if (carril < 0) { finPorCarril.Add(colFin); carril = finPorCarril.Count - 1; }
                else finPorCarril[carril] = colFin;

                semana.Barras.Add(new CalBarra
                {
                    ColInicio = colIni, ColFin = colFin, Carril = carril,
                    Texto = e.Texto, Titulo = e.Titulo, Clase = e.Clase,
                    ContinuaIzquierda = e.Inicio < inicioSemana,
                    ContinuaDerecha = e.Fin > finSemana
                });
            }
            Semanas.Add(semana);
        }
    }
}
