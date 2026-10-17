using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Dashboard;

/// <summary>Lectura gerencial por proyecto. Solo consulta: no sincroniza Squash ni modifica ninguna fuente.</summary>
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public class IndexModel(DashboardService servicio, ILogger<IndexModel> logger) : PageModel
{
    [BindProperty(SupportsGet = true)] public int? InstitucionId { get; set; }
    [BindProperty(SupportsGet = true)] public int? ProyectoId { get; set; }
    [BindProperty(SupportsGet = true)] public int? CampanaId { get; set; }
    [BindProperty(SupportsGet = true)] public int? IteracionId { get; set; }
    [BindProperty(SupportsGet = true)] public int? ModuloId { get; set; }

    public ResultadoDashboard? Resultado { get; private set; }
    public string? ErrorGeneral { get; private set; }

    /// <summary>Con ?actualizar=1 se ignora la caché de Bugzilla y se redirige sin ese parámetro (F5 no repite la consulta forzada).</summary>
    public async Task<IActionResult> OnGetAsync(bool actualizar = false)
    {
        try
        {
            Resultado = await servicio.CargarAsync(new FiltrosDashboard(InstitucionId, ProyectoId, CampanaId, IteracionId, ModuloId), actualizar, HttpContext.RequestAborted);
        }
        catch (OperationCanceledException) when (HttpContext.RequestAborted.IsCancellationRequested) { return new EmptyResult(); }
        catch (Exception ex)
        {
            logger.LogError(ex, "No se pudo armar el Dashboard.");
            ErrorGeneral = "No se pudo leer la información de Portafolio. Intenta de nuevo; si persiste, avisa a un administrador.";
        }
        if (actualizar)
        {
            var filtros = Request.Query.Where(q => !q.Key.Equals("actualizar", StringComparison.OrdinalIgnoreCase))
                .ToDictionary(q => q.Key, q => (string?)q.Value.ToString());
            return RedirectToPage(null, filtros);
        }
        return Page();
    }
}
