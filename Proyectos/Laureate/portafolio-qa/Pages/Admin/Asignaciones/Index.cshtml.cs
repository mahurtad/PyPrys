using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Asignaciones;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public List<Asignacion> Asignaciones { get; set; } = new();

    public async Task OnGetAsync()
    {
        Asignaciones = await _context.Asignaciones
            .Include(a => a.Proyecto)
            .Include(a => a.RecursoQA)
            .OrderByDescending(a => a.FechaInicio)
            .ToListAsync();
    }
}
