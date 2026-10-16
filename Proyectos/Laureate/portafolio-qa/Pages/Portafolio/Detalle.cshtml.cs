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
    public int? DefectosAbiertos => Defectos.Count > 0 && Defectos.All(d => d.Resultado.Disponible)
        ? Defectos.SelectMany(d => d.Resultado.Defectos).DistinctBy(d => d.Id)
            .Count(d => string.IsNullOrWhiteSpace(d.Resolucion)) : null;

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
        foreach (var fuente in proyecto.Integraciones.Where(i => i.Herramienta == HerramientaIntegracion.Bugzilla))
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
