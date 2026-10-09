using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Proyectos;

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
    public Proyecto Proyecto { get; set; } = new();

    [BindProperty]
    public List<int> SistemasSeleccionados { get; set; } = new();

    public List<Institucion> Instituciones { get; set; } = new();
    public List<Sistema> Sistemas { get; set; } = new();

    /// <summary>Recursos QA que se asignan al proyecto al crearlo (hereda las fechas del proyecto).</summary>
    [BindProperty]
    public List<AsignacionInicial> AsignacionesIniciales { get; set; } = new();
    public List<RecursoQA> Recursos { get; set; } = new();

    [BindProperty]
    public bool Forzar { get; set; }

    public List<ConflictoCapacidad> Conflictos { get; set; } = new();
    public bool PuedeForzar => User.IsInRole(Roles.Administrador);

    public async Task<IActionResult> OnGetAsync()
    {
        await CargarCatalogosAsync();

        if (!Instituciones.Any())
        {
            TempData["Mensaje"] = "Registra al menos una institución activa antes de crear un proyecto.";
            return RedirectToPage("/Admin/Instituciones/Index");
        }

        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (Proyecto.FechaInicio > Proyecto.FechaFin)
        {
            ModelState.AddModelError($"{nameof(Proyecto)}.{nameof(Proyecto.FechaFin)}", "La fecha de fin debe ser igual o posterior a la fecha de inicio.");
        }

        // Se ignoran las filas sin recurso seleccionado.
        AsignacionesIniciales = AsignacionesIniciales.Where(a => a.RecursoQAId > 0).ToList();
        foreach (var a in AsignacionesIniciales) a.Rol = string.IsNullOrWhiteSpace(a.Rol) ? "Analista QA" : a.Rol.Trim();

        if (AsignacionesIniciales.Any(a => a.Porcentaje is <= 0 or > 100))
        {
            ModelState.AddModelError(nameof(AsignacionesIniciales), "La dedicación de cada recurso debe estar entre 0.01 y 100.");
        }
        if (AsignacionesIniciales.Any(a => a.Rol.Length > 100))
        {
            ModelState.AddModelError(nameof(AsignacionesIniciales), "El rol no puede superar 100 caracteres.");
        }
        if (AsignacionesIniciales.GroupBy(a => a.RecursoQAId).Any(g => g.Count() > 1))
        {
            ModelState.AddModelError(nameof(AsignacionesIniciales), "Un mismo recurso QA no puede repetirse en el proyecto.");
        }

        var idsValidos = await _context.RecursosQA.Where(r => r.Activo && r.UsuarioId != null).Select(r => r.Id).ToListAsync();
        if (AsignacionesIniciales.Any(a => !idsValidos.Contains(a.RecursoQAId)))
        {
            ModelState.AddModelError(nameof(AsignacionesIniciales), "Todos los recursos deben ser usuarios registrados y activos.");
        }

        if (!ModelState.IsValid)
        {
            await CargarCatalogosAsync();
            return Page();
        }

        Proyecto.Sistemas = SistemasSeleccionados
            .Select(id => new ProyectoSistema { SistemaId = id })
            .ToList();

        Proyecto.Estado = EstadoProyecto.Planificado;
        Proyecto.Avance = 0;
        Proyecto.Actividades = ActividadesPredeterminadas.Crear(Proyecto);

        // Valida la dedicación contra los proyectos vigentes de cada recurso antes de guardar.
        if (Proyecto.Estado is not (EstadoProyecto.Cerrado or EstadoProyecto.Cancelado))
        {
            foreach (var a in AsignacionesIniciales)
            {
                Conflictos.AddRange(await _capacidad.ValidarAsignacionAsync(a.RecursoQAId, Proyecto.FechaInicio, Proyecto.FechaFin, a.Porcentaje));
            }
            if (Conflictos.Count > 0 && !(Forzar && PuedeForzar))
            {
                await CargarCatalogosAsync();
                return Page();
            }
        }

        Proyecto.Asignaciones = AsignacionesIniciales.Select(a => new Asignacion
        {
            RecursoQAId = a.RecursoQAId,
            Rol = a.Rol,
            Porcentaje = a.Porcentaje,
            FechaInicio = Proyecto.FechaInicio.Date,
            FechaFin = Proyecto.FechaFin.Date
        }).ToList();

        _context.Proyectos.Add(Proyecto);
        await _context.SaveChangesAsync();
        var avisos = new List<string>();
        foreach (var a in Proyecto.Asignaciones)
        {
            var resumen = await _capacidad.EvaluarAsync(a.RecursoQAId, a.FechaInicio, a.FechaFin);
            if (resumen is { Sobreasignado: true })
            {
                avisos.Add($"{resumen.Recurso.NombreCompleto} queda sobreasignado/a en {resumen.DiasSobreasignados} día(s) laborable(s)");
            }
        }

        TempData["Mensaje"] = "Proyecto registrado correctamente."
            + (avisos.Count > 0 ? " Atención: " + string.Join("; ", avisos) + "." : "");
        return RedirectToPage("Index");
    }

    private async Task CargarCatalogosAsync()
    {
        Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();
        Sistemas = await _context.Sistemas.Where(s => s.Activo).OrderBy(s => s.Nombre).ToListAsync();
        Recursos = await _context.RecursosQA.Where(r => r.Activo && r.UsuarioId != null).OrderBy(r => r.Nombres).ToListAsync();
    }
}

public class AsignacionInicial
{
    public int RecursoQAId { get; set; }
    public string Rol { get; set; } = "Analista QA";
    public decimal Porcentaje { get; set; } = 100;
}
