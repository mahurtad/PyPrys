using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Portafolio;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
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

    public async Task OnGetAsync()
    {
        Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();
        RecursosQA = await _context.RecursosQA.Where(r => r.Activo).OrderBy(r => r.Nombres).ToListAsync();
        Sistemas = await _context.Sistemas.Where(s => s.Activo).OrderBy(s => s.Nombre).ToListAsync();
        Pms = await _context.Proyectos.Where(p => p.PM != null && p.PM != "").Select(p => p.PM!).Distinct().OrderBy(p => p).ToListAsync();
        Lideres = await _context.Proyectos.Where(p => p.Lider != null && p.Lider != "").Select(p => p.Lider!).Distinct().OrderBy(p => p).ToListAsync();

        var query = _context.Proyectos
            .Include(p => p.Institucion)
            .Include(p => p.Sistemas).ThenInclude(ps => ps.Sistema)
            .Include(p => p.Asignaciones).ThenInclude(a => a.RecursoQA)
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
    }
}
