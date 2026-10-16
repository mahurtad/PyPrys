using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Integraciones;

public class EditModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public EditModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public AsociacionExterna Integracion { get; set; } = new();

    public List<Proyecto> Proyectos { get; set; } = new();

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var integracion = await _context.AsociacionesExternas.FindAsync(id);
        if (integracion is null)
        {
            return NotFound();
        }

        Integracion = integracion;
        await CargarCatalogosAsync(integracion.ProyectoId);
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid)
        {
            await CargarCatalogosAsync(Integracion.ProyectoId);
            return Page();
        }

        _context.Attach(Integracion).State = EntityState.Modified;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Asociación actualizada correctamente.";
        return RedirectToPage("Index");
    }

    private async Task CargarCatalogosAsync(int proyectoActualId)
    {
        Proyectos = await _context.Proyectos
            .Where(p => p.Activo || p.Id == proyectoActualId)
            .OrderBy(p => p.Nombre).ToListAsync();
    }
}
