using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Integraciones;

public class DeleteModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public DeleteModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public AsociacionExterna Integracion { get; set; } = null!;

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var integracion = await _context.AsociacionesExternas
            .Include(a => a.Proyecto)
            .FirstOrDefaultAsync(a => a.Id == id);

        if (integracion is null)
        {
            return NotFound();
        }

        Integracion = integracion;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(int id)
    {
        var integracion = await _context.AsociacionesExternas.FindAsync(id);
        if (integracion is null)
        {
            return NotFound();
        }

        _context.AsociacionesExternas.Remove(integracion);
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Asociación eliminada.";
        return RedirectToPage("Index");
    }
}
