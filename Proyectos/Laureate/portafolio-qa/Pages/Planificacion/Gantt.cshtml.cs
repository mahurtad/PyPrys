using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Planificacion;

public class GanttBarra
{
    public DateTime Inicio { get; set; }
    public DateTime Fin { get; set; }
    public string Texto { get; set; } = "";
    public string Titulo { get; set; } = "";
    public string Clase { get; set; } = "";
    /// <summary>Avance 0-100 a pintar dentro de la barra (solo barras de proyecto).</summary>
    public decimal? Avance { get; set; }
}

public class GanttMarca
{
    public DateTime Fecha { get; set; }
    public string Titulo { get; set; } = "";
}

public class GanttFila
{
    public string Etiqueta { get; set; } = "";
    public string? Subtitulo { get; set; }
    public string? Enlace { get; set; }
    public bool EsDetalle { get; set; }
    public List<GanttBarra> Barras { get; set; } = new();
    public List<GanttMarca> Hitos { get; set; } = new();
}

public class GanttMes
{
    public string Texto { get; set; } = "";
    public double Izquierda { get; set; }
    public double Ancho { get; set; }
}

public class GanttModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly CapacidadService _capacidad;

    public GanttModel(PortafolioQAContext context, CapacidadService capacidad)
    {
        _context = context;
        _capacidad = capacidad;
    }

    public string Vista { get; set; } = "proyecto";
    public DateTime Desde { get; set; }
    public DateTime Hasta { get; set; }
    public int? InstitucionId { get; set; }
    public bool MostrarAsignaciones { get; set; } = true;
    public List<Institucion> Instituciones { get; set; } = new();
    public List<GanttFila> Filas { get; set; } = new();
    public List<GanttMes> Meses { get; set; } = new();
    public int TotalDias { get; set; }
    public double? PosHoy { get; set; }
    public string? Error { get; set; }

    public async Task OnGetAsync(string? vista, DateTime? desde, DateTime? hasta, int? institucionId, bool? asignaciones)
    {
        MostrarAsignaciones = asignaciones ?? true;
        Vista = vista == "recurso" ? "recurso" : "proyecto";
        InstitucionId = institucionId;

        var hoy = DateTime.Today;
        Desde = (desde ?? new DateTime(hoy.Year, hoy.Month, 1).AddMonths(-1)).Date;
        Hasta = (hasta ?? Desde.AddMonths(6).AddDays(-1)).Date;
        if (Hasta < Desde)
        {
            Error = "La fecha final no puede ser anterior a la inicial; se restableció el periodo por defecto.";
            Desde = new DateTime(hoy.Year, hoy.Month, 1).AddMonths(-1);
            Hasta = Desde.AddMonths(6).AddDays(-1);
        }
        if ((Hasta - Desde).Days + 1 > CalculadoraCapacidad.MaximoDiasPeriodo)
        {
            Error = $"El periodo máximo es de {CalculadoraCapacidad.MaximoDiasPeriodo} días; se recortó el fin.";
            Hasta = Desde.AddDays(CalculadoraCapacidad.MaximoDiasPeriodo - 1);
        }

        TotalDias = (Hasta - Desde).Days + 1;
        if (hoy >= Desde && hoy <= Hasta) PosHoy = Posicion(hoy) + Ancho(1) / 2;
        ConstruirMeses();

        Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();

        if (Vista == "proyecto") await CargarPorProyectoAsync();
        else await CargarPorRecursoAsync();
    }

    public double Posicion(DateTime fecha) => (fecha.Date - Desde).Days * 100.0 / TotalDias;
    public double Ancho(int dias) => dias * 100.0 / TotalDias;

    /// <summary>Estilo left/width recortado al periodo visible; null si la barra queda fuera.</summary>
    public string? Estilo(DateTime inicio, DateTime fin)
    {
        var i = inicio.Date < Desde ? Desde : inicio.Date;
        var f = fin.Date > Hasta ? Hasta : fin.Date;
        if (f < i) return null;
        var izq = Posicion(i);
        var ancho = Ancho((f - i).Days + 1);
        return FormattableString.Invariant($"left:{izq:0.####}%;width:{ancho:0.####}%");
    }

    private void ConstruirMeses()
    {
        var cultura = System.Globalization.CultureInfo.GetCultureInfo("es-PE");
        var cursor = Desde;
        while (cursor <= Hasta)
        {
            var finMes = new DateTime(cursor.Year, cursor.Month, 1).AddMonths(1).AddDays(-1);
            if (finMes > Hasta) finMes = Hasta;
            var dias = (finMes - cursor).Days + 1;
            Meses.Add(new GanttMes
            {
                Texto = cursor.ToString("MMM yyyy", cultura),
                Izquierda = Posicion(cursor),
                Ancho = Ancho(dias)
            });
            cursor = finMes.AddDays(1);
        }
    }

    private async Task CargarPorProyectoAsync()
    {
        var consulta = _context.Proyectos.AsNoTracking()
            .Include(p => p.Institucion)
            .Include(p => p.Hitos)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
            .Where(p => p.Activo && p.Estado != EstadoProyecto.Cancelado
                        && p.FechaFin >= Desde && p.FechaInicio <= Hasta);
        if (InstitucionId.HasValue) consulta = consulta.Where(p => p.InstitucionId == InstitucionId);

        var proyectos = await consulta.OrderBy(p => p.FechaInicio).ThenBy(p => p.Nombre).ToListAsync();
        foreach (var p in proyectos)
        {
            Filas.Add(new GanttFila
            {
                Etiqueta = p.Nombre,
                Subtitulo = $"{p.Institucion?.Nombre} · {p.Tipo.Texto()} · {p.Estado.Texto()}",
                Enlace = $"/Portafolio/Detalle/{p.Id}",
                Barras =
                {
                    new GanttBarra
                    {
                        Inicio = p.FechaInicio, Fin = p.FechaFin,
                        Texto = $"{p.Avance:0.#}%",
                        Titulo = $"{p.Nombre}: {p.FechaInicio:dd/MM/yyyy} – {p.FechaFin:dd/MM/yyyy} · avance del proyecto {p.Avance:0.#}%",
                        Clase = "gantt-proyecto", Avance = p.Avance
                    }
                },
                Hitos = p.Hitos.Select(h => new GanttMarca { Fecha = h.Fecha, Titulo = $"{h.Nombre} ({h.Fecha:dd/MM/yyyy})" }).ToList()
            });

            if (!MostrarAsignaciones) continue;

            foreach (var a in p.Asignaciones.Where(x => !x.NoEjecutada).OrderBy(a => a.FechaInicio))
            {
                Filas.Add(new GanttFila
                {
                    Etiqueta = a.RecursoQA?.NombreCompleto ?? "Recurso",
                    Subtitulo = a.Rol,
                    EsDetalle = true,
                    Barras =
                    {
                        new GanttBarra
                        {
                            Inicio = a.FechaInicio, Fin = a.FechaFin,
                            Texto = $"{a.Porcentaje:0.#}%",
                            Titulo = $"{a.RecursoQA?.NombreCompleto} ({a.Rol}): {a.FechaInicio:dd/MM/yyyy} – {a.FechaFin:dd/MM/yyyy} · dedicación {a.Porcentaje:0.#}%",
                            Clase = "gantt-asignacion"
                        }
                    }
                });
            }
        }
    }

    public List<(string Texto, double Izq, double Ancho, string Clase, string Titulo)> Segmentos(ResumenRecurso resumen)
    {
        var lista = new List<(string, double, double, string, string)>();
        DiaCarga? inicioSeg = null;
        DiaCarga? ultimo = null;

        void Cerrar()
        {
            if (inicioSeg is null || ultimo is null) return;
            var dias = (ultimo.Fecha - inicioSeg.Fecha).Days + 1;
            var clase = "seg-" + inicioSeg.Estado.ToString().ToLowerInvariant();
            var titulo = inicioSeg.Estado switch
            {
                EstadoDia.NoLaborable => inicioSeg.MotivoNoLaborable ?? "No laborable",
                EstadoDia.Ausente => "Ausente",
                EstadoDia.Libre => $"Libre · capacidad {inicioSeg.Capacidad:0.#}%",
                _ => $"Carga {inicioSeg.Carga:0.#}% de {inicioSeg.Capacidad:0.#}% de capacidad"
            };
            titulo += $" ({inicioSeg.Fecha:dd/MM} – {ultimo.Fecha:dd/MM})";
            lista.Add((inicioSeg.Estado is EstadoDia.Parcial or EstadoDia.Completo or EstadoDia.Sobreasignado ? $"{inicioSeg.Carga:0.#}%" : "",
                Posicion(inicioSeg.Fecha), Ancho(dias), clase, titulo));
            inicioSeg = null;
        }

        foreach (var dia in resumen.Dias)
        {
            // Fines de semana y feriados se absorben en el segmento vigente para no fragmentar la franja.
            if (!dia.Laborable && inicioSeg is not null)
            {
                ultimo = dia;
                continue;
            }
            if (inicioSeg is not null && (inicioSeg.Estado != dia.Estado || inicioSeg.Carga != dia.Carga))
            {
                Cerrar();
            }
            inicioSeg ??= dia;
            ultimo = dia;
        }
        Cerrar();
        return lista;
    }

    private async Task CargarPorRecursoAsync()
    {
        var resumenes = await _capacidad.CalcularAsync(Desde, Hasta);
        var ids = resumenes.Select(r => r.Recurso.Id).ToList();

        var consulta = _context.Asignaciones.AsNoTracking()
            .Include(a => a.Proyecto).ThenInclude(p => p!.Institucion)
            .Where(a => ids.Contains(a.RecursoQAId) && a.Proyecto!.Estado != EstadoProyecto.Cancelado && !a.NoEjecutada
                        && a.FechaFin >= Desde && a.FechaInicio <= Hasta);
        if (InstitucionId.HasValue) consulta = consulta.Where(a => a.Proyecto!.InstitucionId == InstitucionId);
        var asignaciones = (await consulta.OrderBy(a => a.FechaInicio).ToListAsync()).ToLookup(a => a.RecursoQAId);

        ResumenesPorFila = new();
        foreach (var r in resumenes)
        {
            var fila = new GanttFila
            {
                Etiqueta = r.Recurso.NombreCompleto,
                Subtitulo = $"{r.Recurso.Equipo.Texto()} · capacidad {r.Recurso.CapacidadDisponible:0.#}%"
                            + (r.Sobreasignado ? $" · {r.DiasSobreasignados} d sobreasignado(s)" : "")
            };
            Filas.Add(fila);
            ResumenesPorFila[fila] = r;

            if (!MostrarAsignaciones) continue;

            foreach (var a in asignaciones[r.Recurso.Id])
            {
                Filas.Add(new GanttFila
                {
                    Etiqueta = a.Proyecto?.Nombre ?? "Proyecto",
                    Subtitulo = $"{a.Proyecto?.Institucion?.Nombre} · {a.Rol}",
                    Enlace = $"/Portafolio/Detalle/{a.ProyectoId}",
                    EsDetalle = true,
                    Barras =
                    {
                        new GanttBarra
                        {
                            Inicio = a.FechaInicio, Fin = a.FechaFin,
                            Texto = $"{a.Porcentaje:0.#}%",
                            Titulo = $"{a.Proyecto?.Nombre} ({a.Rol}): {a.FechaInicio:dd/MM/yyyy} – {a.FechaFin:dd/MM/yyyy} · dedicación {a.Porcentaje:0.#}%",
                            Clase = "gantt-asignacion"
                        }
                    }
                });
            }
        }
    }

    public Dictionary<GanttFila, ResumenRecurso> ResumenesPorFila { get; private set; } = new();
}
