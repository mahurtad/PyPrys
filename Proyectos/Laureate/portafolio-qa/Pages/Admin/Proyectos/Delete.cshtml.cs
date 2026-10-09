using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Proyectos;

public class DeleteModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public DeleteModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public Proyecto Proyecto { get; set; } = null!;

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (proyecto is null)
        {
            return NotFound();
        }

        Proyecto = proyecto;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(int id)
    {
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (proyecto is null)
        {
            return NotFound();
        }

        proyecto.Activo = !proyecto.Activo;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = proyecto.Activo ? "Proyecto activado." : "Proyecto desactivado. Se conserva su historial.";
        return RedirectToPage("Index");
    }
}
