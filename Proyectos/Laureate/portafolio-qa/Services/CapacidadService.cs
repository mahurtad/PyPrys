using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Services;

public enum EstadoDia
{
    NoLaborable,
    Ausente,
    Libre,
    Parcial,
    Completo,
    Sobreasignado
}

public class DiaCarga
{
    public DateTime Fecha { get; init; }
    public bool Laborable { get; init; }
    public bool Ausente { get; init; }
    public string? MotivoNoLaborable { get; init; }
    /// <summary>Capacidad del día en %: 0 si no es laborable o la persona está ausente.</summary>
    public decimal Capacidad { get; init; }
    /// <summary>Suma de dedicaciones vigentes ese día. En días no laborables no se contabiliza.</summary>
    public decimal Carga { get; init; }
    public IReadOnlyList<Asignacion> Asignaciones { get; init; } = Array.Empty<Asignacion>();

    public decimal Disponible => Math.Max(0, Capacidad - Carga);
    public decimal Exceso => Math.Max(0, Carga - Capacidad);

    public EstadoDia Estado
    {
        get
        {
            if (!Laborable) return EstadoDia.NoLaborable;
            if (Carga > Capacidad) return EstadoDia.Sobreasignado; // incluye asignado estando ausente
            if (Ausente) return EstadoDia.Ausente;
            if (Carga == 0) return EstadoDia.Libre;
            return Carga == Capacidad ? EstadoDia.Completo : EstadoDia.Parcial;
        }
    }
}

public class RangoSobreasignacion
{
    public DateTime Inicio { get; init; }
    public DateTime Fin { get; init; }
    public decimal CargaMaxima { get; init; }
    public decimal ExcesoMaximo { get; init; }
    public IReadOnlyList<string> Proyectos { get; init; } = Array.Empty<string>();
}

public class ResumenRecurso
{
    public RecursoQA Recurso { get; init; } = null!;
    public IReadOnlyList<DiaCarga> Dias { get; init; } = Array.Empty<DiaCarga>();

    private IEnumerable<DiaCarga> Efectivos => Dias.Where(d => d.Laborable);

    public int DiasLaborables => Efectivos.Count();
    public int DiasSobreasignados => Dias.Count(d => d.Estado == EstadoDia.Sobreasignado);
    public int DiasConDisponibilidad => Efectivos.Count(d => d.Disponible > 0);
    public int DiasAusente => Efectivos.Count(d => d.Ausente);
    public decimal CargaPromedio => DiasLaborables == 0 ? 0 : Math.Round(Efectivos.Average(d => d.Carga), 1);
    public decimal CargaMaxima => DiasLaborables == 0 ? 0 : Efectivos.Max(d => d.Carga);
    public decimal DisponibilidadPromedio => DiasLaborables == 0 ? 0 : Math.Round(Efectivos.Average(d => d.Disponible), 1);
    public bool Sobreasignado => DiasSobreasignados > 0;
    public decimal HorasAsignadas { get; init; }
    public decimal HorasCapacidad { get; init; }
    public IReadOnlyList<RangoSobreasignacion> Rangos { get; init; } = Array.Empty<RangoSobreasignacion>();
}

/// <summary>
/// Cálculo puro de carga/capacidad (sin acceso a datos) para poder verificarlo de forma aislada.
/// Reglas: días laborables = lunes a viernes salvo feriados; en días no laborables no hay capacidad
/// ni carga. Un día de ausencia tiene capacidad 0, de modo que una asignación vigente ese día
/// cuenta como exceso. La carga diaria es la suma de los porcentajes de las asignaciones vigentes.
/// </summary>
public static class CalculadoraCapacidad
{
    public const int MaximoDiasPeriodo = 400;

    public static List<ResumenRecurso> Calcular(
        IEnumerable<RecursoQA> recursos,
        IEnumerable<Asignacion> asignaciones,
        IEnumerable<Feriado> feriados,
        IEnumerable<Ausencia> ausencias,
        DateTime desde,
        DateTime hasta,
        decimal horasPorDia)
    {
        desde = desde.Date;
        hasta = hasta.Date;
        if (hasta < desde) throw new ArgumentException("El fin del periodo no puede ser anterior al inicio.");
        if ((hasta - desde).Days + 1 > MaximoDiasPeriodo)
            throw new ArgumentException("El periodo no puede superar " + MaximoDiasPeriodo + " días.");

        var feriadoPorFecha = feriados.GroupBy(f => f.Fecha.Date).ToDictionary(g => g.Key, g => g.First().Nombre);
        var asigPorRecurso = asignaciones.ToLookup(a => a.RecursoQAId);
        var ausPorRecurso = ausencias.ToLookup(a => a.RecursoQAId);

        var resultado = new List<ResumenRecurso>();
        foreach (var recurso in recursos)
        {
            var asigs = asigPorRecurso[recurso.Id].Where(a => a.FechaFin.Date >= desde && a.FechaInicio.Date <= hasta).ToList();
            var aus = ausPorRecurso[recurso.Id].ToList();
            var dias = new List<DiaCarga>();

            for (var dia = desde; dia <= hasta; dia = dia.AddDays(1))
            {
                string? motivo = null;
                var laborable = dia.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday);
                if (!laborable) motivo = "Fin de semana";
                if (feriadoPorFecha.TryGetValue(dia, out var nombreFeriado))
                {
                    laborable = false;
                    motivo = nombreFeriado;
                }

                var ausente = laborable && aus.Any(a => a.FechaInicio.Date <= dia && dia <= a.FechaFin.Date);
                var vigentes = laborable
                    ? asigs.Where(a => a.FechaInicio.Date <= dia && dia <= a.FechaFin.Date).ToList()
                    : new List<Asignacion>();

                dias.Add(new DiaCarga
                {
                    Fecha = dia,
                    Laborable = laborable,
                    Ausente = ausente,
                    MotivoNoLaborable = motivo,
                    Capacidad = laborable && !ausente ? recurso.CapacidadDisponible : 0,
                    Carga = vigentes.Sum(a => a.Porcentaje),
                    Asignaciones = vigentes
                });
            }

            resultado.Add(new ResumenRecurso
            {
                Recurso = recurso,
                Dias = dias,
                HorasAsignadas = Math.Round(dias.Sum(d => d.Carga) / 100m * horasPorDia, 1),
                HorasCapacidad = Math.Round(dias.Sum(d => d.Capacidad) / 100m * horasPorDia, 1),
                Rangos = ConstruirRangos(dias)
            });
        }

        return resultado;
    }

    // Une días laborables consecutivos con sobreasignación; fines de semana y feriados no cortan el rango.
    private static List<RangoSobreasignacion> ConstruirRangos(List<DiaCarga> dias)
    {
        var rangos = new List<RangoSobreasignacion>();
        List<DiaCarga>? actual = null;

        void Cerrar()
        {
            if (actual is null) return;
            rangos.Add(new RangoSobreasignacion
            {
                Inicio = actual.First().Fecha,
                Fin = actual.Last().Fecha,
                CargaMaxima = actual.Max(d => d.Carga),
                ExcesoMaximo = actual.Max(d => d.Exceso),
                Proyectos = actual.SelectMany(d => d.Asignaciones)
                    .Select(a => a.Proyecto?.Nombre ?? ("Proyecto " + a.ProyectoId))
                    .Distinct().OrderBy(n => n).ToList()
            });
            actual = null;
        }

        foreach (var dia in dias)
        {
            if (!dia.Laborable) continue;
            if (dia.Estado == EstadoDia.Sobreasignado) (actual ??= new()).Add(dia);
            else Cerrar();
        }
        Cerrar();
        return rangos;
    }
}

/// <summary>Tramo de fechas en que una nueva asignación haría que el recurso supere su capacidad.</summary>
public class ConflictoCapacidad
{
    public string Recurso { get; init; } = "";
    public DateTime Inicio { get; init; }
    public DateTime Fin { get; init; }
    public decimal Solicitado { get; init; }
    public decimal Capacidad { get; init; }
    public decimal ExistenteMaximo { get; init; }
    /// <summary>Porcentaje máximo que se puede asignar en TODO el tramo sin sobreasignar.</summary>
    public decimal LibreMinimo { get; init; }
    public bool HayAusencia { get; init; }
    public IReadOnlyList<string> Proyectos { get; init; } = Array.Empty<string>();
}

public class CapacidadService
{
    private readonly PortafolioQAContext _context;
    private readonly decimal _horasPorDia;

    public CapacidadService(PortafolioQAContext context, IConfiguration configuration)
    {
        _context = context;
        _horasPorDia = configuration.GetValue("Planificacion:HorasPorDia", 8m);
    }

    public decimal HorasPorDia => _horasPorDia;

    public async Task<List<ResumenRecurso>> CalcularAsync(
        DateTime desde, DateTime hasta, int? recursoId = null, EquipoQA? equipo = null)
    {
        var recursos = _context.RecursosQA.AsNoTracking().Where(r => r.Activo);
        if (recursoId.HasValue) recursos = recursos.Where(r => r.Id == recursoId);
        if (equipo.HasValue) recursos = recursos.Where(r => r.Equipo == equipo);
        var listaRecursos = await recursos.OrderBy(r => r.Nombres).ThenBy(r => r.Apellidos).ToListAsync();
        var ids = listaRecursos.Select(r => r.Id).ToList();

        var d = desde.Date;
        var h = hasta.Date;
        // Las iniciativas canceladas no consumen capacidad.
        var asignaciones = await _context.Asignaciones.AsNoTracking()
            .Include(a => a.Proyecto)
            .Where(a => ids.Contains(a.RecursoQAId) && a.Proyecto!.Estado != EstadoProyecto.Cancelado && !a.NoEjecutada
                        && a.FechaFin >= d && a.FechaInicio <= h)
            .ToListAsync();
        var feriados = await _context.Feriados.AsNoTracking().Where(f => f.Fecha >= d && f.Fecha <= h).ToListAsync();
        var ausencias = await _context.Ausencias.AsNoTracking()
            .Where(a => ids.Contains(a.RecursoQAId) && a.FechaFin >= d && a.FechaInicio <= h).ToListAsync();

        return CalculadoraCapacidad.Calcular(listaRecursos, asignaciones, feriados, ausencias, d, h, _horasPorDia);
    }

    /// <summary>
    /// Valida ANTES de guardar si asignar <paramref name="porcentaje"/> al recurso en el periodo lo deja por encima de
    /// su capacidad, día por día, sumando las asignaciones vigentes (no canceladas ni no ejecutadas). Solo se evalúan
    /// los días desde hoy: lo ocurrido en el pasado ya no se puede reasignar y no debe impedir planificar hacia adelante. Con
    /// <paramref name="excluirAsignacionId"/> se ignora una asignación existente (edición).
    /// </summary>
    public async Task<List<ConflictoCapacidad>> ValidarAsignacionAsync(
        int recursoId, DateTime desde, DateTime hasta, decimal porcentaje, int? excluirAsignacionId = null)
    {
        var resultado = new List<ConflictoCapacidad>();
        var hoy = DateTime.Today;
        if (desde.Date < hoy) desde = hoy;
        if (hasta.Date < desde.Date) return resultado;
        var resumen = await EvaluarAsync(recursoId, desde, hasta);
        if (resumen is null) return resultado;

        List<(DiaCarga Dia, decimal Existente)>? actual = null;

        void Cerrar()
        {
            if (actual is null) return;
            resultado.Add(new ConflictoCapacidad
            {
                Recurso = resumen.Recurso.NombreCompleto,
                Inicio = actual.First().Dia.Fecha,
                Fin = actual.Last().Dia.Fecha,
                Solicitado = porcentaje,
                Capacidad = resumen.Recurso.CapacidadDisponible,
                ExistenteMaximo = actual.Max(x => x.Existente),
                LibreMinimo = actual.Min(x => Math.Max(0, x.Dia.Capacidad - x.Existente)),
                HayAusencia = actual.Any(x => x.Dia.Ausente),
                Proyectos = actual.SelectMany(x => x.Dia.Asignaciones)
                    .Where(a => a.Id != excluirAsignacionId)
                    .Select(a => a.Proyecto?.Nombre ?? ("Proyecto " + a.ProyectoId))
                    .Distinct().OrderBy(n => n).ToList()
            });
            actual = null;
        }

        foreach (var dia in resumen.Dias)
        {
            if (!dia.Laborable) continue;
            var existente = dia.Carga - dia.Asignaciones.Where(a => a.Id == excluirAsignacionId).Sum(a => a.Porcentaje);
            if (existente + porcentaje > dia.Capacidad) (actual ??= new()).Add((dia, existente));
            else Cerrar();
        }
        Cerrar();
        return resultado;
    }

    /// <summary>Evalúa un recurso en un periodo, para avisar de sobreasignación al guardar una asignación.</summary>
    public async Task<ResumenRecurso?> EvaluarAsync(int recursoId, DateTime desde, DateTime hasta)
    {
        if ((hasta.Date - desde.Date).Days + 1 > CalculadoraCapacidad.MaximoDiasPeriodo)
            hasta = desde.Date.AddDays(CalculadoraCapacidad.MaximoDiasPeriodo - 1);
        var lista = await CalcularAsync(desde, hasta, recursoId);
        return lista.FirstOrDefault();
    }
}
