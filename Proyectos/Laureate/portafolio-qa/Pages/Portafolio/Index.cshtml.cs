using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Pages.Portafolio;

[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly ISquashApi _squash;
    private readonly IConfiguration _config;
    private readonly ILogger<IndexModel> _logger;

    public IndexModel(PortafolioQAContext context, ISquashApi squash, IConfiguration config, ILogger<IndexModel> logger)
    {
        _context = context;
        _squash = squash;
        _config = config;
        _logger = logger;
    }

    [BindProperty(SupportsGet = true)]
    public int? InstitucionId { get; set; }

    [BindProperty(SupportsGet = true)]
    public TipoIniciativa? Tipo { get; set; }

    [BindProperty(SupportsGet = true)]
    public EstadoProyecto? Estado { get; set; }

    [BindProperty(SupportsGet = true)]
    public int? RecursoQAId { get; set; }

    [BindProperty(SupportsGet = true)]
    public string? Pm { get; set; }

    [BindProperty(SupportsGet = true)]
    public string? Lider { get; set; }

    [BindProperty(SupportsGet = true)]
    public int? SistemaId { get; set; }

    public List<Proyecto> Proyectos { get; set; } = new();

    public List<Institucion> Instituciones { get; set; } = new();
    public List<RecursoQA> RecursosQA { get; set; } = new();
    public List<Sistema> Sistemas { get; set; } = new();
    public List<string> Pms { get; set; } = new();
    public List<string> Lideres { get; set; } = new();

    /// <summary>
    /// Al entrar al portafolio se refresca desde Squash TM el avance de los proyectos vinculados cuya última sincronización
    /// es más antigua que Squash:RefrescoMinutos (por defecto 0.5 = 30 s; 0 = siempre). Con ?forzar=1 (botón Actualizar) se ignora el umbral. Si Squash falla o tarda, se muestra lo guardado.
    /// </summary>
    private async Task RefrescarSquashAsync(bool forzar)
    {
        var minutos = _config.GetValue("Squash:RefrescoMinutos", 0.5);
        var limite = forzar ? DateTime.Now.AddDays(1) : DateTime.Now.AddMinutes(-minutos);
        var pendientes = await _context.Proyectos
            .Where(p => p.SquashProjectId != null && p.Estado != EstadoProyecto.Cerrado && p.Estado != EstadoProyecto.Cancelado
                && (p.SquashSincronizadoEn == null || p.SquashSincronizadoEn < limite))
            .ToListAsync();
        if (pendientes.Count == 0) return;

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(HttpContext.RequestAborted);
        cts.CancelAfter(TimeSpan.FromSeconds(30));
        var usuario = User.Identity?.Name ?? CierreProyecto.UsuarioSistema;
        foreach (var proyecto in pendientes)
        {
            try { await SquashSincronizacion.SincronizarAsync(_context, _squash, proyecto, usuario, cts.Token); }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "No se pudo refrescar Squash TM para el proyecto {Id}", proyecto.Id);
                if (cts.IsCancellationRequested) break;
            }
        }
    }

    public async Task<IActionResult> OnGetAsync(bool forzar = false)
    {
        await RefrescarSquashAsync(forzar);
        if (forzar)
        {
            // Redirige sin "forzar" para que recargar con F5 no vuelva a forzar y se conserven los filtros.
            var filtros = Request.Query.Where(q => !q.Key.Equals("forzar", StringComparison.OrdinalIgnoreCase))
                .ToDictionary(q => q.Key, q => (string?)q.Value.ToString());
            return RedirectToPage(null, filtros);
        }

        Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();
        RecursosQA = await _context.RecursosQA.Where(r => r.Activo).OrderBy(r => r.Nombres).ToListAsync();
        Sistemas = await _context.Sistemas.Where(s => s.Activo).OrderBy(s => s.Nombre).ToListAsync();
        Pms = await _context.Proyectos.Where(p => p.PM != null && p.PM != "").Select(p => p.PM!).Distinct().OrderBy(p => p).ToListAsync();
        Lideres = await _context.Proyectos.Where(p => p.Lider != null && p.Lider != "").Select(p => p.Lider!).Distinct().OrderBy(p => p).ToListAsync();

        var query = _context.Proyectos
            .Include(p => p.Institucion)
            .Include(p => p.Sistemas).ThenInclude(ps => ps.Sistema)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
            .Include(p => p.VinculoBugzilla)
            .AsQueryable();

        if (InstitucionId.HasValue)
        {
            query = query.Where(p => p.InstitucionId == InstitucionId.Value);
        }

        if (Tipo.HasValue)
        {
            query = query.Where(p => p.Tipo == Tipo.Value);
        }

        if (Estado.HasValue)
        {
            query = query.Where(p => p.Estado == Estado.Value);
        }

        if (RecursoQAId.HasValue)
        {
            query = query.Where(p => p.Asignaciones.Any(a => a.RecursoQAId == RecursoQAId.Value));
        }

        if (!string.IsNullOrWhiteSpace(Pm))
        {
            query = query.Where(p => p.PM == Pm);
        }

        if (!string.IsNullOrWhiteSpace(Lider))
        {
            query = query.Where(p => p.Lider == Lider);
        }

        if (SistemaId.HasValue)
        {
            query = query.Where(p => p.Sistemas.Any(s => s.SistemaId == SistemaId.Value));
        }

        Proyectos = await query.OrderBy(p => p.Nombre).ToListAsync();
        return Page();
    }
}
