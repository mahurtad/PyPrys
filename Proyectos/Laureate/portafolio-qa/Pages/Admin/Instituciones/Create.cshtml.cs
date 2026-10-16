using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Instituciones;

public class CreateModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public CreateModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Institucion Institucion { get; set; } = new();

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (await _context.Instituciones.AnyAsync(i => i.Nombre == Institucion.Nombre))
        {
            ModelState.AddModelError($"{nameof(Institucion)}.{nameof(Institucion.Nombre)}", "Ya existe una institución con ese nombre.");
        }

        if (!ModelState.IsValid)
        {
            return Page();
        }

        _context.Instituciones.Add(Institucion);
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Institución registrada correctamente.";
        return RedirectToPage("Index");
    }
}
