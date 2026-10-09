using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Instituciones;

public class EditModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public EditModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Institucion Institucion { get; set; } = new();

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

    public async Task<IActionResult> OnPostAsync()
    {
        if (await _context.Instituciones.AnyAsync(i => i.Nombre == Institucion.Nombre && i.Id != Institucion.Id))
        {
            ModelState.AddModelError($"{nameof(Institucion)}.{nameof(Institucion.Nombre)}", "Ya existe una institución con ese nombre.");
        }

        if (!ModelState.IsValid)
        {
            return Page();
        }

        _context.Attach(Institucion).State = EntityState.Modified;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Institución actualizada correctamente.";
        return RedirectToPage("Index");
    }
}
