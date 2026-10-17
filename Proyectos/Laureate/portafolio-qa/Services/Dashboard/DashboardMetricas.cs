using PortafolioQA.Models;
using PortafolioQA.Pages.Portafolio;

namespace PortafolioQA.Services;

/// <summary>Estado de una fuente de datos. Un cero real (Ok con conteo 0) no es lo mismo que "sin datos".</summary>
public enum EstadoFuente
{
    /// <summary>La fuente respondió; un conteo en 0 es un cero real.</summary>
    Ok,
    /// <summary>Hay vínculo pero la fuente aún no tiene información que medir (sin sincronizar, sin casos en el plan).</summary>
    SinDatos,
    /// <summary>El proyecto no está vinculado a un proyecto de Squash o a un producto de Bugzilla.</summary>
    SinVinculo,
    /// <summary>La integración no tiene URL/credencial en la configuración del servidor.</summary>
    NoConfigurada,
    /// <summary>La lectura falló: se omiten las cifras de esa fuente.</summary>
    Error
}

public record EstadoDeFuente(EstadoFuente Estado, string? Detalle = null, DateTime? ActualizadoEn = null)
{
    public string Etiqueta => Estado switch
    {
        EstadoFuente.Ok => "Con datos",
        EstadoFuente.SinDatos => "Sin datos",
        EstadoFuente.SinVinculo => "Sin vínculo",
        EstadoFuente.NoConfigurada => "Integración no configurada",
        _ => "Error"
    };

    public string Badge => Estado switch
    {
        EstadoFuente.Ok => "text-bg-success",
        EstadoFuente.Error => "text-bg-danger",
        EstadoFuente.NoConfigurada => "text-bg-warning",
        _ => "text-bg-secondary"
    };
}

/// <summary>Resultado de pruebas de un conjunto de casos, contando cada caso de prueba una sola vez.</summary>
public record ResumenCertificacion(int Total, int Ejecutados, int Aprobados, int Fallidos, int Bloqueados, int ConAdvertencia, int Excluidos, int EnCurso = 0)
{
    public int Pendientes => Total - Ejecutados;

    /// <summary>Casos que no están ejecutados, bloqueados ni en curso (READY, NOT_RUN y estados desconocidos, como en SquashEstados.Porcentaje).
    /// Aprobados + fallidos + con advertencia + bloqueados + en curso + sin ejecutar = Total.</summary>
    public int SinEjecutar => Total - Ejecutados - Bloqueados - EnCurso;

    /// <summary>Ejecutados / Total. Null si no hay casos (no se muestra 0 %).</summary>
    public decimal? PorcentajeEjecutado => Total == 0 ? null : Math.Round(Ejecutados * 100m / Total, 1);

    /// <summary>Aprobados / Ejecutados. Null si todavía no se ejecutó nada (no se muestra 0 %).</summary>
    public decimal? TasaAprobacion => Ejecutados == 0 ? null : Math.Round(Aprobados * 100m / Ejecutados, 1);
}

public record ResumenPortafolio(int Total, int EnCertificacion, int Vencidos, int ConBloqueantes, int ConBugzillaConsultado,
    IReadOnlyDictionary<EstadoProyecto, int> PorEstado, IReadOnlyDictionary<TipoIniciativa, int> PorTipo)
{
    public int Estado(EstadoProyecto e) => PorEstado.GetValueOrDefault(e);
    public int Tipo(TipoIniciativa t) => PorTipo.GetValueOrDefault(t);
}

/// <summary>Una institución del catálogo con sus proyectos por estado (siempre sobre todo el portafolio, sin importar el filtro).</summary>
public record TarjetaInstitucion(int Id, string Nombre, bool Activa, int Proyectos, IReadOnlyDictionary<EstadoProyecto, int> PorEstado)
{
    public int Estado(EstadoProyecto e) => PorEstado.GetValueOrDefault(e);
}

public record HechoAtencion(string Codigo, string Texto, string Fuente);

public static class DashboardMetricas
{
    // Mismas definiciones que Portafolio → Detalle: Ejecutado = resultado final; Fallido = FAILURE o ERROR.
    private static bool EsAprobado(string e) => e is "SUCCESS" or "SETTLED";
    private static bool EsFallido(string e) => e is "FAILURE" or "ERROR";
    private static bool EsEjecutado(string e) => EsAprobado(e) || EsFallido(e) || e == "WARNING";

    /// <summary>
    /// Total = casos únicos que cuentan para certificar (excluye No ejecutable, Omitido, Cancelado y No encontrado, que se informan
    /// aparte como Excluidos). Un mismo caso de prueba presente en varias iteraciones cuenta una vez, con su último resultado.
    /// Recibe solo casos vigentes del alcance elegido.
    /// </summary>
    public static ResumenCertificacion Certificar(IEnumerable<SquashCaso> casos)
    {
        var lista = casos.ToList();
        var contados = SquashCasosUnicos.Depurar(lista.Where(k => k.Porcentaje is not null));
        var claves = contados.Select(SquashCasosUnicos.Clave).ToHashSet();
        var excluidos = SquashCasosUnicos.Depurar(lista.Where(k => k.Porcentaje is null)).Count(k => !claves.Contains(SquashCasosUnicos.Clave(k)));
        return new ResumenCertificacion(
            contados.Count,
            contados.Count(k => EsEjecutado(k.EstadoSquash)),
            contados.Count(k => EsAprobado(k.EstadoSquash)),
            contados.Count(k => EsFallido(k.EstadoSquash)),
            contados.Count(k => k.EstadoSquash == "BLOCKED"),
            contados.Count(k => k.EstadoSquash == "WARNING"),
            excluidos,
            contados.Count(k => k.EstadoSquash == "RUNNING"));
    }

    /// <summary>
    /// Indicadores de la vista general sobre los proyectos del alcance (todos los de la institución elegida, activos o no, como en Portafolio).
    /// Cuentan proyectos, nunca casos ni defectos, de modo que no hay nada que duplicar. «Con defectos bloqueantes» solo se evalúa
    /// en proyectos cuya consulta a Bugzilla respondió (<see cref="ResumenPortafolio.ConBugzillaConsultado"/>).
    /// </summary>
    public static ResumenPortafolio Resumir(IReadOnlyCollection<ProyectoDashboard> filas) => new(
        filas.Count,
        filas.Count(f => f.Estado == EstadoProyecto.EnCertificacion),
        filas.Count(f => f.DiasVencido is not null),
        filas.Count(f => f.Defectos is { Bloqueantes: > 0 }),
        filas.Count(f => f.Defectos is not null),
        filas.GroupBy(f => f.Estado).ToDictionary(g => g.Key, g => g.Count()),
        filas.GroupBy(f => f.Tipo).ToDictionary(g => g.Key, g => g.Count()));

    public static bool EstaAbierto(EstadoProyecto estado) => estado is not (EstadoProyecto.Cerrado or EstadoProyecto.Cancelado);

    /// <summary>Días de atraso de la fecha comprometida (fecha fin) si el proyecto sigue abierto; null si no está vencido.</summary>
    public static int? DiasVencido(EstadoProyecto estado, DateTime fechaFin, DateTime hoy) =>
        EstaAbierto(estado) && fechaFin.Date < hoy.Date ? (hoy.Date - fechaFin.Date).Days : null;

    /// <summary>
    /// Hechos comprobables, sin ponderar ni resumir en un índice. Cada uno cita su fuente y solo aparece si la fuente
    /// correspondiente respondió (una fuente caída no genera hechos).
    /// </summary>
    public static List<HechoAtencion> Atencion(EstadoProyecto estado, DateTime fechaFin, DateTime hoy,
        ResumenCertificacion? certificacion, DetalleModel.ResumenDefectos? defectos)
    {
        var hechos = new List<HechoAtencion>();
        if (DiasVencido(estado, fechaFin, hoy) is int dias)
            hechos.Add(new("vencido", $"Fecha comprometida vencida hace {dias} día(s) ({fechaFin:dd/MM/yyyy}) con el proyecto aún abierto.", "Portafolio"));
        if (certificacion is { Bloqueados: > 0 })
            hechos.Add(new("pruebas-bloqueadas", $"{certificacion.Bloqueados} caso(s) de prueba bloqueado(s).", "Squash TM"));
        if (defectos is { Bloqueantes: > 0 })
            hechos.Add(new("defectos-bloqueantes", $"{defectos.Bloqueantes} defecto(s) bloqueante(s) abierto(s).", "Bugzilla"));
        if (defectos is { PendienteRetest: > 0 })
            hechos.Add(new("pendiente-retest", $"{defectos.PendienteRetest} defecto(s) corregido(s) pendiente(s) de retest (RESOLVED + FIXED).", "Bugzilla"));
        return hechos;
    }
}
