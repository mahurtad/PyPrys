using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Integraciones;

public class CreateModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public CreateModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public AsociacionExterna Integracion { get; set; } = new();

    public List<Proyecto> Proyectos { get; set; } = new();

    public async Task<IActionResult> OnGetAsync(int? proyectoId)
    {
        await CargarCatalogosAsync();

        if (!Proyectos.Any())
        {
            TempData["Mensaje"] = "Registra al menos un proyecto activo antes de asociar una fuente externa.";
            return RedirectToPage("Index");
        }

        if (proyectoId.HasValue)
        {
            Integracion.ProyectoId = proyectoId.Value;
        }

        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid)
        {
            await CargarCatalogosAsync();
            return Page();
        }

        _context.AsociacionesExternas.Add(Integracion);
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Asociación registrada correctamente.";
        return RedirectToPage("Index");
    }

    private async Task CargarCatalogosAsync()
    {
        Proyectos = await _context.Proyectos.Where(p => p.Activo).OrderBy(p => p.Nombre).ToListAsync();
    }
}
