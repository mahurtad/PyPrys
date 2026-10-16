using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Instituciones;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public List<Institucion> Instituciones { get; set; } = new();

    public async Task OnGetAsync()
    {
        Instituciones = await _context.Instituciones.OrderBy(i => i.Nombre).ToListAsync();
    }
}
