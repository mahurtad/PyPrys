using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Integraciones;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public List<AsociacionExterna> Integraciones { get; set; } = new();

    public async Task OnGetAsync()
    {
        Integraciones = await _context.AsociacionesExternas
            .Include(a => a.Proyecto)
            .OrderBy(a => a.Proyecto!.Nombre)
            .ToListAsync();
    }
}
