using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Pages.Portafolio;

public class DetalleModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly IIntegracionCertificacion _integracionCertificacion;

    public DetalleModel(PortafolioQAContext context, IIntegracionCertificacion integracionCertificacion)
    {
        _context = context;
        _integracionCertificacion = integracionCertificacion;
    }

    public Proyecto Proyecto { get; set; } = null!;
    public List<ResultadoCertificacionConFuente> Certificaciones { get; set; } = new();

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var proyecto = await _context.Proyectos
            .Include(p => p.Institucion)
            .Include(p => p.Sistemas).ThenInclude(ps => ps.Sistema)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
            .Include(p => p.Integraciones)
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

        foreach (var integracion in proyecto.Integraciones.Where(i => i.Herramienta == HerramientaIntegracion.SquashTM))
        {
            var resultado = await _integracionCertificacion.ObtenerAvanceAsync(integracion.IdentificadorExterno);
            Certificaciones.Add(new ResultadoCertificacionConFuente(integracion, resultado));
        }

        return Page();
    }

    public record ResultadoCertificacionConFuente(AsociacionExterna Fuente, ResultadoCertificacion Resultado);
}
