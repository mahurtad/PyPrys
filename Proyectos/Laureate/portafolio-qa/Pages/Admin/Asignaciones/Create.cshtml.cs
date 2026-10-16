using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Asignaciones;

public class CreateModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly CapacidadService _capacidad;

    public CreateModel(PortafolioQAContext context, CapacidadService capacidad)
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

    public async Task<IActionResult> OnGetAsync(int? proyectoId)
    {
        await CargarCatalogosAsync();

        if (!Proyectos.Any() || !Recursos.Any())
        {
            TempData["Mensaje"] = "Registra al menos un proyecto activo y un recurso QA activo antes de crear una asignación.";
            return RedirectToPage("Index");
        }

        if (proyectoId.HasValue)
        {
            Asignacion.ProyectoId = proyectoId.Value;
            var proyecto = Proyectos.FirstOrDefault(p => p.Id == proyectoId.Value);
            if (proyecto is not null)
            {
                // No se precarga el pasado: la planificación nueva empieza hoy como pronto.
                Asignacion.FechaInicio = proyecto.FechaInicio.Date < DateTime.Today ? DateTime.Today : proyecto.FechaInicio;
                Asignacion.FechaFin = proyecto.FechaFin.Date < Asignacion.FechaInicio.Date ? Asignacion.FechaInicio : proyecto.FechaFin;
            }
        }

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
            await CargarCatalogosAsync();
            return Page();
        }

        if (!await _context.RecursosQA.AnyAsync(r => r.Id == Asignacion.RecursoQAId && r.UsuarioId != null))
        {
            ModelState.AddModelError($"{nameof(Asignacion)}.{nameof(Asignacion.RecursoQAId)}", "El recurso debe estar vinculado a un usuario registrado.");
            await CargarCatalogosAsync();
            return Page();
        }

        var estadoDestino = await _context.Proyectos.Where(p => p.Id == Asignacion.ProyectoId).Select(p => p.Estado).FirstOrDefaultAsync();
        if (estadoDestino is EstadoProyecto.Cerrado or EstadoProyecto.Cancelado)
        {
            ModelState.AddModelError($"{nameof(Asignacion)}.{nameof(Asignacion.ProyectoId)}",
                $"El proyecto está {estadoDestino.Texto().ToLowerInvariant()}: no admite nuevas asignaciones. Reábrelo primero si corresponde.");
            await CargarCatalogosAsync();
            return Page();
        }

        if (await ValidarCapacidadAsync())
        {
            await CargarCatalogosAsync();
            return Page();
        }

        _context.Asignaciones.Add(Asignacion);
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Asignación registrada correctamente." + await AvisoSobreasignacionAsync();
        return RedirectToPage("/Portafolio/Detalle", new { id = Asignacion.ProyectoId });
    }

    /// <summary>True si hay sobreasignación que impide guardar (no se forzó o el usuario no es Administrador).</summary>
    private async Task<bool> ValidarCapacidadAsync()
    {
        var estado = await _context.Proyectos.Where(p => p.Id == Asignacion.ProyectoId).Select(p => p.Estado).FirstOrDefaultAsync();
        if (estado is EstadoProyecto.Cerrado or EstadoProyecto.Cancelado) return false;

        Conflictos = await _capacidad.ValidarAsignacionAsync(Asignacion.RecursoQAId, Asignacion.FechaInicio, Asignacion.FechaFin, Asignacion.Porcentaje);
        return Conflictos.Count > 0 && !(Forzar && PuedeForzar);
    }

    private async Task CargarCatalogosAsync()
    {
        Proyectos = await _context.Proyectos.Where(p => p.Activo).OrderBy(p => p.Nombre).ToListAsync();
        Recursos = await _context.RecursosQA.Where(r => r.Activo && r.UsuarioId != null).OrderBy(r => r.Nombres).ToListAsync();
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
