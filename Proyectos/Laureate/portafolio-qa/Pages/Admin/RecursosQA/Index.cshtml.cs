using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.RecursosQA;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public List<RecursoQA> Recursos { get; set; } = new();

    public async Task OnGetAsync()
    {
        Recursos = await _context.RecursosQA.OrderBy(r => r.Nombres).ToListAsync();
    }
}
