using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Ausencias;

public class IndexModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public IndexModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Ausencia Nueva { get; set; } = new();

    public List<Ausencia> Ausencias { get; set; } = new();
    public List<RecursoQA> Recursos { get; set; } = new();

    public async Task OnGetAsync() => await CargarAsync();

    public async Task<IActionResult> OnPostAsync()
    {
        if (Nueva.FechaInicio > Nueva.FechaFin)
        {
            ModelState.AddModelError("Nueva.FechaFin", "La fecha de fin debe ser igual o posterior a la fecha de inicio.");
        }

        if (!ModelState.IsValid)
        {
            await CargarAsync();
            return Page();
        }

        Nueva.FechaInicio = Nueva.FechaInicio.Date;
        Nueva.FechaFin = Nueva.FechaFin.Date;
        Nueva.Motivo = string.IsNullOrWhiteSpace(Nueva.Motivo) ? null : Nueva.Motivo.Trim();
        _context.Ausencias.Add(Nueva);
        await _context.SaveChangesAsync();
        TempData["Mensaje"] = "Ausencia registrada.";
        return RedirectToPage();
    }

    public async Task<IActionResult> OnPostEliminarAsync(int id)
    {
        var ausencia = await _context.Ausencias.FindAsync(id);
        if (ausencia is not null)
        {
            _context.Ausencias.Remove(ausencia);
            await _context.SaveChangesAsync();
            TempData["Mensaje"] = "Ausencia eliminada.";
        }
        return RedirectToPage();
    }

    private async Task CargarAsync()
    {
        Recursos = await _context.RecursosQA.Where(r => r.Activo).OrderBy(r => r.Nombres).ToListAsync();
        Ausencias = await _context.Ausencias.Include(a => a.RecursoQA)
            .OrderByDescending(a => a.FechaInicio).ToListAsync();
    }
}
