using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Feriados;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Feriado Nuevo { get; set; } = new();

    public List<Feriado> Feriados { get; set; } = new();

    public async Task OnGetAsync() => await CargarAsync();

    public async Task<IActionResult> OnPostAsync()
    {
        Nuevo.Nombre = Nuevo.Nombre?.Trim() ?? string.Empty;
        if (ModelState.IsValid && await _context.Feriados.AnyAsync(f => f.Fecha == Nuevo.Fecha.Date))
        {
            ModelState.AddModelError("Nuevo.Fecha", "Ya existe un feriado en esa fecha.");
        }

        if (!ModelState.IsValid)
        {
            await CargarAsync();
            return Page();
        }

        Nuevo.Fecha = Nuevo.Fecha.Date;
        _context.Feriados.Add(Nuevo);
        await _context.SaveChangesAsync();
        TempData["Mensaje"] = "Feriado registrado.";
        return RedirectToPage();
    }

    public async Task<IActionResult> OnPostEliminarAsync(int id)
    {
        var feriado = await _context.Feriados.FindAsync(id);
        if (feriado is not null)
        {
            _context.Feriados.Remove(feriado);
            await _context.SaveChangesAsync();
            TempData["Mensaje"] = "Feriado eliminado.";
        }
        return RedirectToPage();
    }

    private async Task CargarAsync() =>
        Feriados = await _context.Feriados.OrderByDescending(f => f.Fecha).ToListAsync();
}
