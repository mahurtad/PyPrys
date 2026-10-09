using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Asignaciones;

public class DeleteModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public DeleteModel(PortafolioQAContext context)
    {
        _context = context;
    }

    public Asignacion Asignacion { get; set; } = null!;

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var asignacion = await _context.Asignaciones
            .Include(a => a.Proyecto)
            .Include(a => a.RecursoQA)
            .FirstOrDefaultAsync(a => a.Id == id);

        if (asignacion is null)
        {
            return NotFound();
        }

        Asignacion = asignacion;
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(int id)
    {
        var asignacion = await _context.Asignaciones.FindAsync(id);
        if (asignacion is null)
        {
            return NotFound();
        }

        var proyectoId = asignacion.ProyectoId;
        var recursoId = asignacion.RecursoQAId;
        if (await PortafolioQA.Services.CierreProyecto.EquipoBloqueadoAsync(_context, proyectoId))
        {
            TempData["Error"] = "El proyecto está cerrado: sus asignaciones son historial y no se pueden eliminar. Reábrelo primero si corresponde.";
            return RedirectToPage("/Portafolio/Detalle", new { id = proyectoId });
        }
        _context.Asignaciones.Remove(asignacion);
        await _context.SaveChangesAsync();
        var liberadas = await PortafolioQA.Services.CierreProyecto.DesvincularActividadesHuerfanasAsync(_context, proyectoId, recursoId);

        TempData["Mensaje"] = "Asignación eliminada."
            + (liberadas > 0 ? $" {liberadas} actividad(es) quedaron sin recurso asignado." : "");
        return RedirectToPage("/Portafolio/Detalle", new { id = proyectoId });
    }
}
