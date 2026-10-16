using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Sistemas;

public class EditModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public EditModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Sistema Sistema { get; set; } = new();

    [BindProperty]
    public List<int> InstitucionesSeleccionadas { get; set; } = new();

    public List<Institucion> Instituciones { get; set; } = new();

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var sistema = await _context.Sistemas
            .Include(s => s.Instituciones)
            .FirstOrDefaultAsync(s => s.Id == id);

        if (sistema is null)
        {
            return NotFound();
        }

        Sistema = sistema;
        InstitucionesSeleccionadas = sistema.Instituciones.Select(si => si.InstitucionId).ToList();
        Instituciones = await _context.Instituciones.Where(i => i.Activo || InstitucionesSeleccionadas.Contains(i.Id))
            .OrderBy(i => i.Nombre).ToListAsync();
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (await _context.Sistemas.AnyAsync(s => s.Nombre == Sistema.Nombre && s.Id != Sistema.Id))
        {
            ModelState.AddModelError($"{nameof(Sistema)}.{nameof(Sistema.Nombre)}", "Ya existe un sistema con ese nombre.");
        }

        if (!ModelState.IsValid)
        {
            Instituciones = await _context.Instituciones.OrderBy(i => i.Nombre).ToListAsync();
            return Page();
        }

        var sistema = await _context.Sistemas
            .Include(s => s.Instituciones)
            .FirstOrDefaultAsync(s => s.Id == Sistema.Id);

        if (sistema is null)
        {
            return NotFound();
        }

        sistema.Nombre = Sistema.Nombre;
        sistema.Activo = Sistema.Activo;
        sistema.Instituciones.Clear();
        foreach (var institucionId in InstitucionesSeleccionadas)
        {
            sistema.Instituciones.Add(new SistemaInstitucion { SistemaId = sistema.Id, InstitucionId = institucionId });
        }

        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Sistema actualizado correctamente.";
        return RedirectToPage("Index");
    }
}
