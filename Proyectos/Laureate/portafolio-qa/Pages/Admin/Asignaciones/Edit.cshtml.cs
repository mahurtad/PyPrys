using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Asignaciones;

public class EditModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly CapacidadService _capacidad;

    public EditModel(PortafolioQAContext context, CapacidadService capacidad)
    {
        _context = context;
        _capacidad = capacidad;
    }

    [BindProperty]
    public Asignacion Asignacion { get; set; } = new();

    [BindProperty]
    public bool Forzar { get; set; }

    public List<ConflictoCapacidad> Conflictos { get; set; } = new();
    public bool PuedeForzar => User.IsInRole(Roles.Administrador);

    public List<Proyecto> Proyectos { get; set; } = new();
    public List<RecursoQA> Recursos { get; set; } = new();

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var asignacion = await _context.Asignaciones.FindAsync(id);
        if (asignacion is null)
        {
            return NotFound();
        }

        if (await PortafolioQA.Services.CierreProyecto.EquipoBloqueadoAsync(_context, asignacion.ProyectoId))
        {
            TempData["Error"] = "El proyecto está cerrado: sus asignaciones son historial y no se pueden editar. Reábrelo primero si corresponde.";
            return RedirectToPage("/Portafolio/Detalle", new { id = asignacion.ProyectoId });
        }

        Asignacion = asignacion;
        await CargarCatalogosAsync(asignacion.ProyectoId, asignacion.RecursoQAId);
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (Asignacion.FechaInicio > Asignacion.FechaFin)
        {
            ModelState.AddModelError($"{nameof(Asignacion)}.{nameof(Asignacion.FechaFin)}", "La fecha de fin debe ser igual o posterior a la fecha de inicio.");
        }

        if (!ModelState.IsValid)
        {
            await CargarCatalogosAsync(Asignacion.ProyectoId, Asignacion.RecursoQAId);
            return Page();
        }

        var original = await _context.Asignaciones.AsNoTracking().Where(a => a.Id == Asignacion.Id).Select(a => a.RecursoQAId).FirstOrDefaultAsync();
        if (Asignacion.RecursoQAId != original
            && !await _context.RecursosQA.AnyAsync(r => r.Id == Asignacion.RecursoQAId && r.UsuarioId != null))
        {
            ModelState.AddModelError($"{nameof(Asignacion)}.{nameof(Asignacion.RecursoQAId)}", "El recurso debe estar vinculado a un usuario registrado.");
            await CargarCatalogosAsync(Asignacion.ProyectoId, original);
            return Page();
        }

        var proyectoOriginal = await _context.Asignaciones.AsNoTracking().Where(a => a.Id == Asignacion.Id).Select(a => a.ProyectoId).FirstOrDefaultAsync();
        var estadoDestino = await _context.Proyectos.Where(p => p.Id == Asignacion.ProyectoId).Select(p => p.Estado).FirstOrDefaultAsync();
        if (await PortafolioQA.Services.CierreProyecto.EquipoBloqueadoAsync(_context, proyectoOriginal)
            || estadoDestino is EstadoProyecto.Cerrado or EstadoProyecto.Cancelado)
        {
            ModelState.AddModelError($"{nameof(Asignacion)}.{nameof(Asignacion.ProyectoId)}",
                "El proyecto (de origen o destino) está cerrado o cancelado: no admite cambios en su equipo. Reábrelo primero si corresponde.");
            await CargarCatalogosAsync(Asignacion.ProyectoId, Asignacion.RecursoQAId);
            return Page();
        }

        if (await ValidarCapacidadAsync())
        {
            await CargarCatalogosAsync(Asignacion.ProyectoId, Asignacion.RecursoQAId);
            return Page();
        }

        var anterior = await _context.Asignaciones.AsNoTracking().Where(a => a.Id == Asignacion.Id)
            .Select(a => new { a.ProyectoId, a.RecursoQAId }).FirstOrDefaultAsync();

        _context.Attach(Asignacion).State = EntityState.Modified;
        await _context.SaveChangesAsync();

        var liberadas = anterior is not null && (anterior.ProyectoId != Asignacion.ProyectoId || anterior.RecursoQAId != Asignacion.RecursoQAId)
            ? await PortafolioQA.Services.CierreProyecto.DesvincularActividadesHuerfanasAsync(_context, anterior.ProyectoId, anterior.RecursoQAId)
            : 0;

        TempData["Mensaje"] = "Asignación actualizada correctamente." + await AvisoSobreasignacionAsync()
            + (liberadas > 0 ? $" {liberadas} actividad(es) quedaron sin recurso asignado." : "");
        return RedirectToPage("/Portafolio/Detalle", new { id = Asignacion.ProyectoId });
    }

    /// <summary>True si hay sobreasignación que impide guardar (no se forzó o el usuario no es Administrador).</summary>
    private async Task<bool> ValidarCapacidadAsync()
    {
        var estado = await _context.Proyectos.Where(p => p.Id == Asignacion.ProyectoId).Select(p => p.Estado).FirstOrDefaultAsync();
        if (estado is EstadoProyecto.Cerrado or EstadoProyecto.Cancelado) return false;

        Conflictos = await _capacidad.ValidarAsignacionAsync(Asignacion.RecursoQAId, Asignacion.FechaInicio, Asignacion.FechaFin, Asignacion.Porcentaje, Asignacion.Id);
        return Conflictos.Count > 0 && !(Forzar && PuedeForzar);
    }

    private async Task CargarCatalogosAsync(int proyectoActualId, int recursoActualId)
    {
        Proyectos = await _context.Proyectos
            .Where(p => p.Activo || p.Id == proyectoActualId)
            .OrderBy(p => p.Nombre).ToListAsync();
        Recursos = await _context.RecursosQA
            .Where(r => (r.Activo && r.UsuarioId != null) || r.Id == recursoActualId)
            .OrderBy(r => r.Nombres).ToListAsync();
    }

    private async Task<string> AvisoSobreasignacionAsync()
    {
        var resumen = await _capacidad.EvaluarAsync(Asignacion.RecursoQAId, Asignacion.FechaInicio, Asignacion.FechaFin);
        if (resumen is null || !resumen.Sobreasignado)
        {
            return string.Empty;
        }

        var rangos = string.Join("; ", resumen.Rangos.Take(3).Select(r =>
            $"{r.Inicio:dd/MM/yyyy}-{r.Fin:dd/MM/yyyy} (pico {r.CargaMaxima:0.#}%)"));
        return $" Atención: {resumen.Recurso.NombreCompleto} queda sobreasignado/a en {resumen.DiasSobreasignados} día(s) laborable(s): {rangos}.";
    }
}
