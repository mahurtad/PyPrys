using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Sistemas;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public List<Sistema> Sistemas { get; set; } = new();

    public async Task OnGetAsync()
    {
        Sistemas = await _context.Sistemas
            .Include(s => s.Instituciones).ThenInclude(si => si.Institucion)
            .OrderBy(s => s.Nombre)
            .ToListAsync();
    }
}
