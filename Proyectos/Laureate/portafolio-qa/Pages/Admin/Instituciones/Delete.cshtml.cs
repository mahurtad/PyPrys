using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Instituciones;

/// <summary>
/// No elimina el registro: alterna su estado activo para conservar el
/// historial de proyectos y sistemas relacionados.
/// </summary>
public class DeleteModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public DeleteModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public Institucion Institucion { get; set; } = null!;

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var institucion = await _context.Instituciones.FindAsync(id);
        if (institucion is null)
        {
            return NotFound();
        }

        Institucion = institucion;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(int id)
    {
        var institucion = await _context.Instituciones.FindAsync(id);
        if (institucion is null)
        {
            return NotFound();
        }

        institucion.Activo = !institucion.Activo;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = institucion.Activo ? "Institución activada." : "Institución desactivada. Se conserva su historial.";
        return RedirectToPage("Index");
    }
}
