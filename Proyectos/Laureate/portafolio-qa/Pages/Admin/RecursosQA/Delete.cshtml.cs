using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.RecursosQA;

public class DeleteModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public DeleteModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public RecursoQA Recurso { get; set; } = null!;

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var recurso = await _context.RecursosQA.FindAsync(id);
        if (recurso is null)
        {
            return NotFound();
        }

        Recurso = recurso;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(int id)
    {
        var recurso = await _context.RecursosQA.FindAsync(id);
        if (recurso is null)
        {
            return NotFound();
        }

        recurso.Activo = !recurso.Activo;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = recurso.Activo ? "Recurso activado." : "Recurso desactivado. Se conservan sus asignaciones.";
        return RedirectToPage("Index");
    }
}
