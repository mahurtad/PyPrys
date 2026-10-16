using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Proyectos;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public List<Proyecto> Proyectos { get; set; } = new();
    public List<Institucion> Instituciones { get; set; } = new();
    public int Total { get; set; }

    [BindProperty(SupportsGet = true)] public string? Q { get; set; }
    [BindProperty(SupportsGet = true)] public EstadoProyecto? Estado { get; set; }
    [BindProperty(SupportsGet = true)] public int? InstitucionId { get; set; }
    [BindProperty(SupportsGet = true)] public bool VerInactivos { get; set; }
    public bool HayFiltros => !string.IsNullOrWhiteSpace(Q) || Estado.HasValue || InstitucionId.HasValue || VerInactivos;

    public async Task OnGetAsync()
    {
        Total = await _context.Proyectos.CountAsync();
        Instituciones = await _context.Instituciones.OrderBy(i => i.Nombre).ToListAsync();

        var consulta = _context.Proyectos
            .Include(p => p.Institucion)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
            .AsQueryable();
        if (!VerInactivos) consulta = consulta.Where(p => p.Activo);
        if (Estado.HasValue) consulta = consulta.Where(p => p.Estado == Estado.Value);
        if (InstitucionId.HasValue) consulta = consulta.Where(p => p.InstitucionId == InstitucionId.Value);
        if (!string.IsNullOrWhiteSpace(Q))
        {
            var q = Q.Trim();
            consulta = consulta.Where(p => p.Nombre.Contains(q) || (p.PM != null && p.PM.Contains(q)) || (p.Lider != null && p.Lider.Contains(q)));
        }
        Proyectos = await consulta.OrderBy(p => p.Nombre).ToListAsync();
    }
}
