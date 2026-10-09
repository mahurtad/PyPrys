using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Sistemas;

public class DeleteModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public DeleteModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public Sistema Sistema { get; set; } = null!;

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var sistema = await _context.Sistemas.FindAsync(id);
        if (sistema is null)
        {
            return NotFound();
        }

        Sistema = sistema;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(int id)
    {
        var sistema = await _context.Sistemas.FindAsync(id);
        if (sistema is null)
        {
            return NotFound();
        }

        sistema.Activo = !sistema.Activo;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = sistema.Activo ? "Sistema activado." : "Sistema desactivado. Se conserva su historial.";
        return RedirectToPage("Index");
    }
}
