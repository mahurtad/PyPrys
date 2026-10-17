using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Pages.Portafolio;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Services;

public record FiltrosDashboard(int? InstitucionId, int? ProyectoId, int? CampanaId, int? IteracionId, int? ModuloId);

public record MiembroEquipo(string Nombre, string Rol, decimal Dedicacion, DateTime Inicio, DateTime Fin, bool VigenteHoy);
public record NodoModulo(int Id, string Nombre, ResumenCertificacion Certificacion);
public record NodoIteracion(int Id, string Nombre, string? Estado, ResumenCertificacion Certificacion, int CasosSinModulo, List<NodoModulo> Modulos);
public record NodoCampana(int Id, string Nombre, string? Estado, ResumenCertificacion Certificacion, List<NodoIteracion> Iteraciones);

/// <summary>Defecto de Bugzilla que requiere atención (bloqueante abierto o pendiente de retest).</summary>
public record DefectoAtencion(string Id, string Titulo, string Estado, string? Severidad, string? Responsable, string? Url);

public class ProyectoDashboard
{
    public int Id { get; init; }
    public string Nombre { get; init; } = "";
    public int InstitucionId { get; init; }
    public string Institucion { get; init; } = "";
    public TipoIniciativa Tipo { get; init; }
    public EstadoProyecto Estado { get; init; }
    public bool Activo { get; init; }
    public string? Bpo { get; init; }
    public string? Pm { get; init; }
    public string? Lider { get; init; }
    public DateTime FechaInicio { get; init; }
    public DateTime FechaFin { get; init; }
    /// <summary>Avance del proyecto tal como lo guarda Portafolio (actividades o promedio de ejecución en Squash). No es % de casos ejecutados.</summary>
    public decimal Avance { get; init; }
    public bool Abierto { get; init; }
    public int? DiasVencido { get; init; }
    public List<MiembroEquipo> Equipo { get; init; } = new();

    public EstadoDeFuente Squash { get; set; } = new(EstadoFuente.SinVinculo);
    public ResumenCertificacion? Certificacion { get; set; }

    public EstadoDeFuente Bugzilla { get; set; } = new(EstadoFuente.SinVinculo);
    public DetalleModel.ResumenDefectos? Defectos { get; set; }
    public List<string> Productos { get; init; } = new();
    public List<DefectoAtencion> DefectosBloqueantes { get; set; } = new();
    public List<DefectoAtencion> DefectosRetest { get; set; } = new();

    public List<HechoAtencion> Atencion { get; set; } = new();
    /// <summary>Fuentes cuyas cifras faltan o no son confiables; el resumen se muestra igual con lo disponible.</summary>
    public List<string> Incompleto { get; } = new();
}

public class DetalleSeleccion
{
    public required ProyectoDashboard Proyecto { get; init; }
    public int? CampanaId { get; init; }
    public int? IteracionId { get; init; }
    public int? ModuloId { get; init; }
    public string? EtiquetaAlcance { get; init; }
    public List<NodoCampana> Campanas { get; init; } = new();
    /// <summary>Todas las campañas del proyecto (para llenar los filtros, sin importar el alcance elegido).</summary>
    public List<NodoCampana> TodasLasCampanas { get; init; } = new();
}

public class ResultadoDashboard
{
    public DateTime ConsultadoEn { get; init; }
    public FiltrosDashboard Filtros { get; init; } = new(null, null, null, null, null);
    public List<Institucion> Instituciones { get; init; } = new();
    public List<(int Id, string Nombre, int InstitucionId)> Proyectos { get; init; } = new();
    public List<ProyectoDashboard> Filas { get; init; } = new();
    /// <summary>Una tarjeta por institución, sobre todo el portafolio (no cambia con el filtro de institución).</summary>
    public List<TarjetaInstitucion> TarjetasInstituciones { get; init; } = new();
    /// <summary>Indicadores de las filas mostradas (el alcance de todos los gráficos, la tabla y las alertas).</summary>
    public ResumenPortafolio Resumen { get; set; } = DashboardMetricas.Resumir(Array.Empty<ProyectoDashboard>());
    public DetalleSeleccion? Detalle { get; set; }
    public EstadoDeFuente FuentePortafolio { get; set; } = new(EstadoFuente.Ok);
    public EstadoDeFuente FuenteSquash { get; set; } = new(EstadoFuente.SinVinculo);
    public EstadoDeFuente FuenteBugzilla { get; set; } = new(EstadoFuente.SinVinculo);
    public DateTime? SquashMasAntiguo { get; set; }
    public DateTime? SquashMasReciente { get; set; }
    public DateTime? BugzillaMasAntiguo { get; set; }
    public DateTime? BugzillaMasReciente { get; set; }
    public List<string> Avisos { get; } = new();
}

/// <summary>
/// Caché breve de las consultas de defectos a Bugzilla por nombre de producto, con un máximo de consultas simultáneas.
/// Singleton: lo comparten todas las visitas al Dashboard. Los fallos se guardan menos tiempo que los aciertos.
/// </summary>
public class DashboardDefectosCache(IMemoryCache cache)
{
    public static readonly TimeSpan VidaAcierto = TimeSpan.FromSeconds(60);
    public static readonly TimeSpan VidaFallo = TimeSpan.FromSeconds(15);
    private const int MaximoSimultaneas = 4;
    private readonly SemaphoreSlim _cupo = new(MaximoSimultaneas);

    private static string Clave(string producto) => "dashboard:bugzilla:" + producto.Trim().ToLowerInvariant();

    public void Invalidar(string producto) => cache.Remove(Clave(producto));

    private static async Task<bool?> SinErrorAsync(Func<Task<bool?>> accion)
    {
        try { return await accion(); } catch (Exception ex) when (ex is not OperationCanceledException) { return null; }
    }

    public async Task<ResultadoDefectos> ObtenerAsync(IIntegracionDefectos bugzilla, string producto, bool forzar, CancellationToken ct)
    {
        var clave = Clave(producto);
        if (forzar) cache.Remove(clave);
        if (cache.TryGetValue(clave, out ResultadoDefectos? guardado) && guardado is not null) return guardado;

        await _cupo.WaitAsync(ct);
        try
        {
            if (cache.TryGetValue(clave, out guardado) && guardado is not null) return guardado;
            ResultadoDefectos r;
            try { r = await bugzilla.ObtenerDefectosAsync(producto, ct); }
            catch (Exception ex) when (ex is not OperationCanceledException)
            { r = new ResultadoDefectos { Disponible = false, Mensaje = "No se pudo consultar Bugzilla." }; }
            // Cero defectos solo es un cero real si el producto existe: Bugzilla no falla ante un nombre inexistente o sin permiso.
            if (r.Disponible && r.Defectos.Count == 0 && bugzilla is IVerificaProductoDefectos verifica
                && await SinErrorAsync(() => verifica.ProductoExisteAsync(producto, ct)) == false)
                r = new ResultadoDefectos { Disponible = false, Mensaje = "El producto no existe en Bugzilla o la cuenta no tiene acceso. Revisa su nombre." };
            r.UltimaActualizacion ??= DateTime.UtcNow;
            cache.Set(clave, r, r.Disponible ? VidaAcierto : VidaFallo);
            return r;
        }
        finally { _cupo.Release(); }
    }
}

/// <summary>
/// Arma el Dashboard gerencial. Solo lee: Portafolio y los datos de Squash ya sincronizados (sin llamar a Squash ni escribir)
/// y los defectos de Bugzilla (GET, con caché breve). Cada fuente falla por separado.
/// </summary>
public class DashboardService(PortafolioQAContext db, IIntegracionDefectos bugzilla, DashboardDefectosCache cacheDefectos,
    IOptions<SquashOptions> squashOpciones, IOptions<BugzillaOptions> bugzillaOpciones, TimeProvider reloj)
{
    private sealed record CasoCarga(int Id, int IteracionId, int ItemId, int? TestCaseId, string Estado, decimal? Porcentaje, DateTime? Ultima);

    public async Task<ResultadoDashboard> CargarAsync(FiltrosDashboard filtros, bool forzar, CancellationToken ct)
    {
        var ahora = reloj.GetLocalNow().DateTime;
        var hoy = ahora.Date;

        var catalogo = await db.Proyectos.AsNoTracking().OrderBy(p => p.Nombre).Select(p => new { p.Id, p.Nombre, p.InstitucionId, p.Estado }).ToListAsync(ct);
        var opciones = catalogo.Select(p => (p.Id, p.Nombre, p.InstitucionId)).ToList();
        // Instituciones activas y las inactivas que aún tienen proyectos (si no, esos proyectos no sumarían en ninguna tarjeta).
        var conProyectos = catalogo.Select(p => p.InstitucionId).ToHashSet();
        var instituciones = (await db.Instituciones.AsNoTracking().OrderBy(i => i.Nombre).ToListAsync(ct))
            .Where(i => i.Activo || conProyectos.Contains(i.Id)).ToList();
        var tarjetas = instituciones.Select(i =>
        {
            var propios = catalogo.Where(p => p.InstitucionId == i.Id).ToList();
            return new TarjetaInstitucion(i.Id, i.Nombre, i.Activo, propios.Count, propios.GroupBy(p => p.Estado).ToDictionary(g => g.Key, g => g.Count()));
        }).ToList();

        // Un proyecto elegido manda sobre la institución: la selección nunca queda incoherente.
        var seleccionado = filtros.ProyectoId is int pid ? catalogo.FirstOrDefault(p => p.Id == pid) : null;
        var institucionId = seleccionado?.InstitucionId ?? filtros.InstitucionId;
        var resultado = new ResultadoDashboard
        {
            ConsultadoEn = ahora,
            Filtros = filtros with { InstitucionId = institucionId, ProyectoId = seleccionado?.Id },
            Instituciones = instituciones,
            TarjetasInstituciones = tarjetas,
            Proyectos = opciones
        };
        if (filtros.ProyectoId is not null && seleccionado is null) resultado.Avisos.Add("El proyecto indicado no existe; se ignoró ese filtro.");

        var proyectos = await db.Proyectos.AsNoTracking()
            .Include(p => p.Institucion)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
            .Include(p => p.Integraciones)
            .Include(p => p.VinculoBugzilla)
            .Where(p => institucionId == null || p.InstitucionId == institucionId)
            .OrderBy(p => p.Nombre)
            .AsSplitQuery()
            .ToListAsync(ct);
        var ids = proyectos.Select(p => p.Id).ToList();

        foreach (var p in proyectos)
        {
            var pd = new ProyectoDashboard
            {
                Id = p.Id, Nombre = p.Nombre, InstitucionId = p.InstitucionId, Institucion = p.Institucion?.Nombre ?? "—",
                Tipo = p.Tipo, Estado = p.Estado, Activo = p.Activo, Bpo = p.BPO, Pm = p.PM, Lider = p.Lider,
                FechaInicio = p.FechaInicio, FechaFin = p.FechaFin, Avance = p.Avance,
                Abierto = DashboardMetricas.EstaAbierto(p.Estado), DiasVencido = DashboardMetricas.DiasVencido(p.Estado, p.FechaFin, hoy),
                Equipo = p.Asignaciones.Where(a => !a.NoEjecutada).OrderBy(a => a.FechaInicio)
                    .Select(a => new MiembroEquipo(a.RecursoQA?.NombreCompleto ?? "—", a.Rol, a.Porcentaje, a.FechaInicio, a.FechaFin,
                        a.FechaInicio.Date <= hoy && a.FechaFin.Date >= hoy)).ToList()
            };
            foreach (var producto in ProductosDe(p)) pd.Productos.Add(producto);
            resultado.Filas.Add(pd);
        }

        // --- Squash TM: datos sincronizados ---
        List<SquashCaso> casos = new();
        List<(int Id, int ProyectoId, int SquashId, string Nombre, string? Estado, List<(int Id, string Nombre, string? Estado, int SquashId)> Iteraciones)> campanas = new();
        List<(int Id, int IteracionId, string Nombre, List<int> CasoIds)> modulos = new();
        string? errorSquash = null;
        try
        {
            var conSquash = proyectos.Where(p => p.SquashProjectId != null).Select(p => p.Id).ToList();
            if (conSquash.Count > 0)
            {
                var cs = await db.SquashCampanas.AsNoTracking().Where(c => c.Vigente && conSquash.Contains(c.ProyectoId))
                    .Select(c => new { c.Id, c.ProyectoId, c.SquashId, c.Nombre, c.Estado,
                        Its = c.Iteraciones.Where(i => i.Vigente).Select(i => new { i.Id, i.Nombre, i.Estado, i.SquashId }).ToList() })
                    .ToListAsync(ct);
                campanas = cs.Select(c => (c.Id, c.ProyectoId, c.SquashId, c.Nombre, c.Estado,
                    c.Its.Select(i => (i.Id, i.Nombre, i.Estado, i.SquashId)).ToList())).ToList();
                var ks = await db.SquashCasos.AsNoTracking()
                    .Where(k => k.Vigente && k.SquashIteracion!.Vigente && k.SquashIteracion.SquashCampana!.Vigente && conSquash.Contains(k.SquashIteracion.SquashCampana.ProyectoId))
                    .Select(k => new CasoCarga(k.Id, k.SquashIteracionId, k.SquashItemId, k.SquashTestCaseId, k.EstadoSquash, k.Porcentaje, k.UltimaEjecucion))
                    .ToListAsync(ct);
                casos = ks.Select(k => new SquashCaso { Id = k.Id, SquashIteracionId = k.IteracionId, SquashItemId = k.ItemId, SquashTestCaseId = k.TestCaseId,
                    EstadoSquash = k.Estado, Porcentaje = k.Porcentaje, UltimaEjecucion = k.Ultima }).ToList();
                var ms = await db.SquashModulos.AsNoTracking()
                    .Where(m => m.Vigente && m.SquashIteracion!.Vigente && m.SquashIteracion.SquashCampana!.Vigente && conSquash.Contains(m.SquashIteracion.SquashCampana.ProyectoId))
                    .Select(m => new { m.Id, m.SquashIteracionId, m.Nombre, Casos = m.Casos.Select(k => k.Id).ToList() })
                    .ToListAsync(ct);
                modulos = ms.Select(m => (m.Id, m.SquashIteracionId, m.Nombre, m.Casos)).ToList();
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            errorSquash = "No se pudieron leer los datos sincronizados de Squash TM.";
            casos = new(); campanas = new(); modulos = new();
        }

        var casosPorIteracion = casos.ToLookup(k => k.SquashIteracionId);
        foreach (var p in proyectos)
        {
            var pd = resultado.Filas.First(f => f.Id == p.Id);
            AplicarSquash(pd, p, errorSquash, campanas.Where(c => c.ProyectoId == p.Id).ToList(), casosPorIteracion, ahora);
        }

        // --- Bugzilla: una consulta por producto distinto, en paralelo y con cupo ---
        var productos = resultado.Filas.SelectMany(f => f.Productos).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var respuestas = new Dictionary<string, ResultadoDefectos>(StringComparer.OrdinalIgnoreCase);
        if (bugzillaOpciones.Value.Configurado && productos.Count > 0)
        {
            var tareas = productos.Select(async prod => (prod, r: await cacheDefectos.ObtenerAsync(bugzilla, prod, forzar, ct))).ToList();
            foreach (var (prod, r) in await Task.WhenAll(tareas)) respuestas[prod] = r;
        }
        foreach (var p in proyectos)
            AplicarBugzilla(resultado.Filas.First(f => f.Id == p.Id), p, respuestas, hoy);

        // Atención a nivel de fila (alcance completo del proyecto).
        foreach (var f in resultado.Filas)
            f.Atencion = DashboardMetricas.Atencion(f.Estado, f.FechaFin, hoy, f.Certificacion, f.Defectos);

        ResumirFuentes(resultado);
        resultado.Resumen = DashboardMetricas.Resumir(resultado.Filas);

        // --- Detalle del proyecto elegido, con el alcance de campaña / iteración / módulo ---
        if (seleccionado is not null)
        {
            var fila = resultado.Filas.First(f => f.Id == seleccionado.Id);
            resultado.Detalle = ArmarDetalle(fila, filtros, campanas.Where(c => c.ProyectoId == fila.Id).ToList(), casosPorIteracion, modulos, resultado, hoy);
        }
        else if (filtros.CampanaId != null || filtros.IteracionId != null || filtros.ModuloId != null)
            resultado.Avisos.Add("Los filtros de campaña, iteración y módulo requieren elegir un proyecto; se ignoraron.");

        return resultado;
    }

    /// <summary>Productos de Bugzilla del proyecto: asociaciones manuales y producto preparado por el vínculo (mismo criterio que Detalle).</summary>
    private static IEnumerable<string> ProductosDe(Proyecto p)
    {
        var lista = p.Integraciones.Where(i => i.Herramienta == HerramientaIntegracion.Bugzilla && !string.IsNullOrWhiteSpace(i.IdentificadorExterno))
            .Select(i => i.IdentificadorExterno.Trim()).ToList();
        var vinculado = p.VinculoBugzilla?.ProductoNombre?.Trim();
        if (!string.IsNullOrWhiteSpace(vinculado) && !lista.Any(x => x.Equals(vinculado, StringComparison.OrdinalIgnoreCase))) lista.Add(vinculado);
        return lista.Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private void AplicarSquash(ProyectoDashboard pd, Proyecto p, string? errorSquash,
        List<(int Id, int ProyectoId, int SquashId, string Nombre, string? Estado, List<(int Id, string Nombre, string? Estado, int SquashId)> Iteraciones)> campanas,
        ILookup<int, SquashCaso> casosPorIteracion, DateTime ahora)
    {
        if (p.SquashProjectId is null)
        {
            pd.Squash = new(EstadoFuente.SinVinculo, "El proyecto no está vinculado a un proyecto de Squash TM.");
            return;
        }
        var sincronizado = p.SquashSincronizadoEn;
        if (errorSquash is not null)
        {
            pd.Squash = new(EstadoFuente.Error, errorSquash, sincronizado);
            pd.Incompleto.Add("Squash TM: " + errorSquash);
            return;
        }
        if (sincronizado is null)
        {
            pd.Squash = new(EstadoFuente.SinDatos, "Vinculado, pero todavía no se sincronizó.");
            return;
        }
        var delProyecto = campanas.SelectMany(c => c.Iteraciones).SelectMany(i => casosPorIteracion[i.Id]).ToList();
        var cert = DashboardMetricas.Certificar(delProyecto);
        if (!squashOpciones.Value.Configurado)
        {
            // Se conservan los datos de la última sincronización, pero se avisa que no se pueden refrescar.
            pd.Squash = new(EstadoFuente.NoConfigurada, "Integración no configurada: se muestran los datos de la última sincronización.", sincronizado);
            pd.Incompleto.Add("Squash TM: integración no configurada; los datos pueden estar desactualizados.");
            pd.Certificacion = cert.Total == 0 ? null : cert;
            return;
        }
        if (cert.Total == 0)
        {
            pd.Squash = new(EstadoFuente.SinDatos, cert.Excluidos > 0 ? "Solo hay casos excluidos del conteo (no ejecutable, omitido o cancelado)." : "El plan de pruebas no tiene casos.", sincronizado);
            return;
        }
        pd.Squash = new(EstadoFuente.Ok, null, sincronizado);
        pd.Certificacion = cert;
    }

    private void AplicarBugzilla(ProyectoDashboard pd, Proyecto p, Dictionary<string, ResultadoDefectos> respuestas, DateTime hoy)
    {
        if (pd.Productos.Count == 0)
        {
            var v = p.VinculoBugzilla;
            if (v?.Estado == EstadoBugzilla.Error)
            {
                pd.Bugzilla = new(EstadoFuente.Error, "La preparación del producto en Bugzilla falló; reintenta desde el detalle del proyecto.");
                pd.Incompleto.Add("Bugzilla: la preparación del producto falló.");
            }
            else
                pd.Bugzilla = new(EstadoFuente.SinVinculo, v is null ? "El proyecto no tiene producto de Bugzilla asociado." : "Producto de Bugzilla pendiente de preparar.");
            return;
        }
        if (!bugzillaOpciones.Value.Configurado)
        {
            pd.Bugzilla = new(EstadoFuente.NoConfigurada, "Falta la URL o la clave de API de Bugzilla en la configuración del servidor.");
            pd.Incompleto.Add("Bugzilla: integración no configurada.");
            return;
        }
        var resultados = pd.Productos.Select(prod => respuestas.TryGetValue(prod, out var r) ? r : null).ToList();
        var fecha = resultados.Where(r => r?.UltimaActualizacion != null).Select(r => (DateTime?)r!.UltimaActualizacion!.Value.ToLocalTime()).OrderBy(f => f).FirstOrDefault();
        var fallos = resultados.Where(r => r is null || !r.Disponible).ToList();
        if (fallos.Count > 0)
        {
            // Con varios productos, un resultado parcial subestimaría los defectos: no se muestran cifras (igual que Detalle).
            var motivo = fallos.Select(r => r?.Mensaje).FirstOrDefault(m => !string.IsNullOrWhiteSpace(m)) ?? "No se pudo consultar Bugzilla.";
            pd.Bugzilla = new(EstadoFuente.Error, motivo, fecha);
            pd.Incompleto.Add("Bugzilla: " + motivo);
            return;
        }
        var defectos = resultados.SelectMany(r => r!.Defectos).DistinctBy(d => d.Id).ToList();
        pd.Defectos = DetalleModel.ResumenDefectos.De(defectos);
        // Cero real: Bugzilla respondió y el producto no tiene defectos.
        pd.Bugzilla = new(EstadoFuente.Ok, defectos.Count == 0 ? "El producto no tiene defectos registrados." : null, fecha);
        DefectoAtencion Mapear(DefectoExterno d) => new(d.Id, d.Titulo, d.EstadoTexto, d.Severidad, d.Responsable, d.Url);
        pd.DefectosBloqueantes = defectos.Where(d => DetalleModel.ResumenDefectos.De(new[] { d }).Bloqueantes == 1).Select(Mapear).Take(10).ToList();
        pd.DefectosRetest = defectos.Where(d => DetalleModel.ResumenDefectos.De(new[] { d }).PendienteRetest == 1).Select(Mapear).Take(10).ToList();
    }

    private static void ResumirFuentes(ResultadoDashboard r)
    {
        r.FuentePortafolio = new(EstadoFuente.Ok, null, r.ConsultadoEn);
        var sq = r.Filas.Where(f => f.Squash.ActualizadoEn != null).Select(f => f.Squash.ActualizadoEn!.Value).ToList();
        if (sq.Count > 0) { r.SquashMasAntiguo = sq.Min(); r.SquashMasReciente = sq.Max(); }
        var bz = r.Filas.Where(f => f.Bugzilla.ActualizadoEn != null).Select(f => f.Bugzilla.ActualizadoEn!.Value).ToList();
        if (bz.Count > 0) { r.BugzillaMasAntiguo = bz.Min(); r.BugzillaMasReciente = bz.Max(); }
        r.FuenteSquash = ResumenDeEstados(r.Filas.Select(f => f.Squash.Estado));
        r.FuenteBugzilla = ResumenDeEstados(r.Filas.Select(f => f.Bugzilla.Estado));
    }

    /// <summary>Estado global de una fuente: Error si alguna falló; si no, Ok si alguna tiene datos; si no, el estado predominante.</summary>
    private static EstadoDeFuente ResumenDeEstados(IEnumerable<EstadoFuente> estados)
    {
        var l = estados.ToList();
        if (l.Count == 0) return new(EstadoFuente.SinVinculo);
        if (l.Contains(EstadoFuente.Error)) return new(EstadoFuente.Error, $"{l.Count(e => e == EstadoFuente.Error)} proyecto(s) con error.");
        if (l.Contains(EstadoFuente.Ok)) return new(EstadoFuente.Ok);
        if (l.Contains(EstadoFuente.NoConfigurada)) return new(EstadoFuente.NoConfigurada);
        if (l.Contains(EstadoFuente.SinDatos)) return new(EstadoFuente.SinDatos);
        return new(EstadoFuente.SinVinculo);
    }

    private DetalleSeleccion ArmarDetalle(ProyectoDashboard fila, FiltrosDashboard f,
        List<(int Id, int ProyectoId, int SquashId, string Nombre, string? Estado, List<(int Id, string Nombre, string? Estado, int SquashId)> Iteraciones)> campanas,
        ILookup<int, SquashCaso> casosPorIteracion, List<(int Id, int IteracionId, string Nombre, List<int> CasoIds)> modulos,
        ResultadoDashboard resultado, DateTime hoy)
    {
        var todas = new List<NodoCampana>();
        foreach (var c in campanas)
        {
            var its = new List<NodoIteracion>();
            foreach (var i in c.Iteraciones)
            {
                var casosIt = casosPorIteracion[i.Id].ToList();
                var mods = modulos.Where(m => m.IteracionId == i.Id)
                    .Select(m => new NodoModulo(m.Id, m.Nombre, DashboardMetricas.Certificar(casosIt.Where(k => m.CasoIds.Contains(k.Id))))).ToList();
                var enAlgunModulo = modulos.Where(m => m.IteracionId == i.Id).SelectMany(m => m.CasoIds).ToHashSet();
                // Casos del plan que ninguna suite de Squash contiene: se informan como "sin módulo", no se reparten.
                var sinModulo = DashboardMetricas.Certificar(casosIt.Where(k => !enAlgunModulo.Contains(k.Id))).Total;
                its.Add(new NodoIteracion(i.Id, i.Nombre, i.Estado, DashboardMetricas.Certificar(casosIt), sinModulo, mods));
            }
            todas.Add(new NodoCampana(c.Id, c.Nombre, c.Estado, DashboardMetricas.Certificar(c.Iteraciones.SelectMany(i => casosPorIteracion[i.Id])), its));
        }

        // El filtro más específico manda; los demás se deducen de la jerarquía real (nunca por nombres).
        int? mid = null, iid = null, cid = null;
        string? etiqueta = null;
        if (f.ModuloId is int mf)
        {
            var par = todas.SelectMany(c => c.Iteraciones.Select(i => (c, i))).FirstOrDefault(x => x.i.Modulos.Any(m => m.Id == mf));
            if (par.i is null) resultado.Avisos.Add("El módulo indicado no pertenece al proyecto; se ignoró ese filtro.");
            else { mid = mf; iid = par.i.Id; cid = par.c.Id; etiqueta = $"Módulo «{par.i.Modulos.First(x => x.Id == mf).Nombre}» · iteración «{par.i.Nombre}» · campaña «{par.c.Nombre}»"; }
        }
        if (mid is null && f.IteracionId is int itf)
        {
            var par = todas.SelectMany(c => c.Iteraciones.Select(i => (c, i))).FirstOrDefault(x => x.i.Id == itf);
            if (par.i is null) resultado.Avisos.Add("La iteración indicada no pertenece al proyecto; se ignoró ese filtro.");
            else { iid = itf; cid = par.c.Id; etiqueta = $"Iteración «{par.i.Nombre}» · campaña «{par.c.Nombre}»"; }
        }
        if (iid is null && f.CampanaId is int cf)
        {
            var c = todas.FirstOrDefault(x => x.Id == cf);
            if (c is null) resultado.Avisos.Add("La campaña indicada no pertenece al proyecto; se ignoró ese filtro.");
            else { cid = cf; etiqueta = $"Campaña «{c.Nombre}»"; }
        }
        if ((f.CampanaId is int c0 && cid is not null && c0 != cid) || (f.IteracionId is int i0 && iid is not null && i0 != iid))
            resultado.Avisos.Add("Los filtros de campaña, iteración y módulo no eran coherentes entre sí; se aplicó el más específico.");

        // Certificación del alcance elegido (los casos se vuelven a depurar dentro de ese alcance).
        var alcance = campanas.Where(c => cid == null || c.Id == cid).SelectMany(c => c.Iteraciones).Where(i => iid == null || i.Id == iid)
            .SelectMany(i => casosPorIteracion[i.Id]).ToList();
        if (mid is int mm) { var ids = modulos.First(x => x.Id == mm).CasoIds.ToHashSet(); alcance = alcance.Where(k => ids.Contains(k.Id)).ToList(); }
        var certAlcance = DashboardMetricas.Certificar(alcance);

        // Hechos de atención con la certificación del alcance; defectos y vencimiento no dependen de campaña/iteración/módulo.
        var detalle = new ProyectoDashboard
        {
            Id = fila.Id, Nombre = fila.Nombre, InstitucionId = fila.InstitucionId, Institucion = fila.Institucion, Tipo = fila.Tipo, Estado = fila.Estado,
            Activo = fila.Activo, Bpo = fila.Bpo, Pm = fila.Pm, Lider = fila.Lider, FechaInicio = fila.FechaInicio, FechaFin = fila.FechaFin,
            Avance = fila.Avance, Abierto = fila.Abierto, DiasVencido = fila.DiasVencido, Equipo = fila.Equipo, Productos = fila.Productos,
            Squash = fila.Squash, Bugzilla = fila.Bugzilla, Defectos = fila.Defectos,
            DefectosBloqueantes = fila.DefectosBloqueantes, DefectosRetest = fila.DefectosRetest
        };
        detalle.Certificacion = fila.Certificacion is null ? null : certAlcance;
        if (fila.Squash.Estado == EstadoFuente.Ok && certAlcance.Total == 0)
            detalle.Squash = new(EstadoFuente.SinDatos, "La selección no tiene casos que cuenten para certificar.", fila.Squash.ActualizadoEn);
        foreach (var inc in fila.Incompleto) detalle.Incompleto.Add(inc);
        detalle.Atencion = DashboardMetricas.Atencion(fila.Estado, fila.FechaFin, hoy, detalle.Certificacion, detalle.Defectos);

        var visibles = todas.Where(c => cid == null || c.Id == cid)
            .Select(c => c with { Iteraciones = c.Iteraciones.Where(i => iid == null || i.Id == iid)
                .Select(i => i with { Modulos = i.Modulos.Where(m => mid == null || m.Id == mid).ToList() }).ToList() }).ToList();

        return new DetalleSeleccion { Proyecto = detalle, CampanaId = cid, IteracionId = iid, ModuloId = mid, EtiquetaAlcance = etiqueta, Campanas = visibles, TodasLasCampanas = todas };
    }
}
