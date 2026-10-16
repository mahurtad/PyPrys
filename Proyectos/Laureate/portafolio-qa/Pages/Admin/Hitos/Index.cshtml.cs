using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Hitos;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Hito Nuevo { get; set; } = new();

    public List<Hito> Hitos { get; set; } = new();
    public List<Proyecto> Proyectos { get; set; } = new();

    public async Task OnGetAsync(int? proyectoId)
    {
        if (proyectoId.HasValue) Nuevo.ProyectoId = proyectoId.Value;
        await CargarAsync();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        Nuevo.Nombre = Nuevo.Nombre?.Trim() ?? string.Empty;
        if (!ModelState.IsValid)
        {
            await CargarAsync();
            return Page();
        }

        Nuevo.Fecha = Nuevo.Fecha.Date;
        _context.Hitos.Add(Nuevo);
        await _context.SaveChangesAsync();
        TempData["Mensaje"] = "Hito registrado.";
        return RedirectToPage();
    }

    public async Task<IActionResult> OnPostEliminarAsync(int id)
    {
        var hito = await _context.Hitos.FindAsync(id);
        if (hito is not null)
        {
            _context.Hitos.Remove(hito);
            await _context.SaveChangesAsync();
            TempData["Mensaje"] = "Hito eliminado.";
        }
        return RedirectToPage();
    }

    private async Task CargarAsync()
    {
        Proyectos = await _context.Proyectos.Where(p => p.Activo).OrderBy(p => p.Nombre).ToListAsync();
        Hitos = await _context.Hitos.Include(h => h.Proyecto).OrderByDescending(h => h.Fecha).ToListAsync();
    }
}
