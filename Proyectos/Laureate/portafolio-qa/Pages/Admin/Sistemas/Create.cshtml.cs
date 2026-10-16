using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Sistemas;

public class CreateModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public CreateModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Sistema Sistema { get; set; } = new();

    [BindProperty]
    public List<int> InstitucionesSeleccionadas { get; set; } = new();

    public List<Institucion> Instituciones { get; set; } = new();

    public async Task OnGetAsync()
    {
        Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (await _context.Sistemas.AnyAsync(s => s.Nombre == Sistema.Nombre))
        {
            ModelState.AddModelError($"{nameof(Sistema)}.{nameof(Sistema.Nombre)}", "Ya existe un sistema con ese nombre.");
        }

        if (!ModelState.IsValid)
        {
            Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();
            return Page();
        }

        Sistema.Instituciones = InstitucionesSeleccionadas
            .Select(id => new SistemaInstitucion { InstitucionId = id })
            .ToList();

        _context.Sistemas.Add(Sistema);
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Sistema registrado correctamente.";
        return RedirectToPage("Index");
    }
}
