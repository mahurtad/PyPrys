using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Pages.Portafolio;

[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public class DetalleModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly IIntegracionCertificacion _integracionCertificacion;
    private readonly ISquashApi _squash;
    private readonly IIntegracionDefectos _defectos;
    private readonly BugzillaProvision _bugzilla;

    public DetalleModel(PortafolioQAContext context, IIntegracionCertificacion integracionCertificacion, ISquashApi squash, IIntegracionDefectos defectos, BugzillaProvision bugzilla)
    {
        _bugzilla = bugzilla;
        _context = context;
        _integracionCertificacion = integracionCertificacion;
        _squash = squash;
        _defectos = defectos;
    }

    public Proyecto Proyecto { get; set; } = null!;
    public List<ResultadoCertificacionConFuente> Certificaciones { get; set; } = new();
    public StatusCertificacion Status { get; private set; } = new(0, 0, 0);
    public List<(AsociacionExterna Fuente, ResultadoDefectos Resultado)> Defectos { get; } = new();
    /// <summary>Resumen de defectos con las mismas reglas del dashboard Kiwi/Bugzilla; null si Bugzilla no está disponible.</summary>
    public ResumenDefectos? Resumen => Defectos.Count > 0 && Defectos.All(d => d.Resultado.Disponible)
        ? ResumenDefectos.De(Defectos.SelectMany(d => d.Resultado.Defectos).DistinctBy(d => d.Id)) : null;
    public int? DefectosAbiertos => Resumen?.Abiertos;

    /// <summary>Abierto = estado distinto de RESOLVED, VERIFIED y CLOSED. Pendiente de retest = RESOLVED + FIXED.</summary>
    public record ResumenDefectos(int Total, int Abiertos, int Bloqueantes, int NoBloqueantes, int Mejoras, int SinClasificar, int PendienteRetest)
    {
        public static ResumenDefectos De(IEnumerable<DefectoExterno> defectos)
        {
            var lista = defectos.ToList();
            var abiertos = lista.Where(d => !new[] { "RESOLVED", "VERIFIED", "CLOSED" }.Contains(d.Estado.Trim().ToUpperInvariant())).ToList();
            int Con(string sev) => abiertos.Count(d => Normalizar(d.Severidad) == sev);
            int b = Con("bloqueante"), nb = Con("no-bloqueante"), m = Con("mejora");
            var retest = lista.Count(d => d.Estado.Trim().Equals("RESOLVED", StringComparison.OrdinalIgnoreCase)
                && string.Equals(d.Resolucion?.Trim(), "FIXED", StringComparison.OrdinalIgnoreCase));
            return new ResumenDefectos(lista.Count, abiertos.Count, b, nb, m, abiertos.Count - b - nb - m, retest);
        }

        private static string Normalizar(string? valor)
        {
            var d = (valor ?? "").ToLowerInvariant().Normalize(System.Text.NormalizationForm.FormD);
            var sinTildes = new string(d.Where(c => System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark).ToArray());
            var token = System.Text.RegularExpressions.Regex.Replace(sinTildes, "[^a-z0-9]+", "-").Trim('-');
            // Severidades estándar de Bugzilla (inglés) equivalentes a las del dashboard en español.
            return token switch
            {
                "blocker" => "bloqueante",
                "enhancement" => "mejora",
                "critical" or "major" or "normal" or "minor" or "trivial" => "no-bloqueante",
                _ => token
            };
        }
    }

    /// <summary>Casos vigentes que cuentan en la certificación (excluye No ejecutable, Omitido y Cancelado).</summary>
    public record StatusCertificacion(int Total, int Ejecutados, int Fallidos);

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var proyecto = await _context.Proyectos
            .Include(p => p.Institucion)
            .Include(p => p.Sistemas).ThenInclude(ps => ps.Sistema)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
            .Include(p => p.Integraciones)
            .Include(p => p.VinculoBugzilla)
            .Include(p => p.SquashCampanas).ThenInclude(c => c.Iteraciones).ThenInclude(i => i.Casos).ThenInclude(k => k.RecursoQA)
            .Include(p => p.SquashCampanas).ThenInclude(c => c.Iteraciones).ThenInclude(i => i.Modulos).ThenInclude(m => m.Casos)
            .Include(p => p.Acuerdos).ThenInclude(x => x.RecursoQA)
            .Include(p => p.Acuerdos).ThenInclude(x => x.Actividades).ThenInclude(v => v.Actividad)
            .Include(p => p.Actividades).ThenInclude(a => a.RecursoQA)
            .Include(p => p.Cierres).ThenInclude(c => c.Participantes)
            .Include(p => p.Cierres).ThenInclude(c => c.Actividades)
            .AsSplitQuery()
            .FirstOrDefaultAsync(p => p.Id == id);

        if (proyecto is null)
        {
            return NotFound();
        }

        Proyecto = proyecto;
        // Fuentes de defectos: asociaciones manuales (Integraciones) y el producto preparado automáticamente (vínculo Bugzilla).
        var fuentes = proyecto.Integraciones.Where(i => i.Herramienta == HerramientaIntegracion.Bugzilla).ToList();
        var productoVinculado = proyecto.VinculoBugzilla?.ProductoNombre;
        if (!string.IsNullOrWhiteSpace(productoVinculado)
            && !fuentes.Any(f => f.IdentificadorExterno.Trim().Equals(productoVinculado.Trim(), StringComparison.OrdinalIgnoreCase)))
            fuentes.Add(new AsociacionExterna { ProyectoId = proyecto.Id, Herramienta = HerramientaIntegracion.Bugzilla, IdentificadorExterno = productoVinculado });
        foreach (var fuente in fuentes)
            Defectos.Add((fuente, await _defectos.ObtenerDefectosAsync(fuente.IdentificadorExterno, HttpContext.RequestAborted)));

        // Ejecutado = con resultado final (Aprobado, Corregido, Fallido, Error, Advertencia). Fallido = FAILURE o ERROR.
        // Cada caso de prueba cuenta una sola vez (resultado más reciente), aunque esté en varias iteraciones.
        var casos = SquashCasosUnicos.Depurar(proyecto.SquashCampanas.Where(c => c.Vigente)
            .SelectMany(c => c.Iteraciones.Where(i => i.Vigente))
            .SelectMany(i => i.Casos.Where(k => k.Vigente && k.Porcentaje is not null)));
        Status = new StatusCertificacion(
            casos.Count,
            casos.Count(k => k.EstadoSquash is "SUCCESS" or "SETTLED" or "FAILURE" or "ERROR" or "WARNING"),
            casos.Count(k => k.EstadoSquash is "FAILURE" or "ERROR"));

        // Squash ya no se consulta al abrir la página: el avance llega por sincronización (Edit → "Actualizar desde Squash")
        // y la sección Certificación está oculta. Al reactivarla, cargar aquí bajo demanda.

        return Page();
    }

    public async Task<IActionResult> OnPostSincronizarSquashAsync(int id)
    {
        // [Authorize] no aplica a handlers de Razor Pages: la restricción a Administrador/Editor se valida aquí.
        if (!User.IsInRole(Roles.Administrador) && !User.IsInRole(Roles.Editor)) return Forbid();
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (proyecto is null) return NotFound();
        TempData["Mensaje"] = await SquashSincronizacion.SincronizarAsync(_context, _squash, proyecto, User.Identity?.Name ?? CierreProyecto.UsuarioSistema);
        return RedirectToPage(null, null, new { id }, "squash");
    }

    /// <summary>Reintenta desde el paso pendiente. "asociarExistente" (solo administrador) confirma un producto preexistente con el mismo nombre.</summary>
    public async Task<IActionResult> OnPostReintentarBugzillaAsync(int id, bool asociarExistente = false)
    {
        if (!User.IsInRole(Roles.Administrador) && !User.IsInRole(Roles.Editor)) return Forbid();
        if (asociarExistente && !User.IsInRole(Roles.Administrador)) return Forbid();
        var r = await _bugzilla.PrepararAsync(id, asociarExistente, HttpContext.RequestAborted);
        if (r.Estado == EstadoBugzilla.Configurado) TempData["Mensaje"] = r.Mensaje;
        else TempData["Error"] = "Bugzilla: " + r.Mensaje;
        return RedirectToPage(null, null, new { id }, "bugzilla");
    }

    public record ResultadoCertificacionConFuente(AsociacionExterna Fuente, ResultadoCertificacion Resultado);
}
