using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Pages.Admin.Proyectos;

public class CreateModel : PageModel
{
    private readonly PortafolioQAContext _context;
    private readonly CapacidadService _capacidad;
    private readonly ISquashApi _squash;
    private readonly BugzillaProvision _bugzilla;

    public CreateModel(PortafolioQAContext context, CapacidadService capacidad, ISquashApi squash, BugzillaProvision bugzilla)
    {
        _bugzilla = bugzilla;
        _context = context;
        _capacidad = capacidad;
        _squash = squash;
    }

    /// <summary>Todos los proyectos de Squash TM (para validar).</summary>
    public List<SquashProyectoDto> SquashProyectos { get; set; } = new();
    /// <summary>Los que aún no están vinculados a un proyecto del portafolio: lo único que se ofrece al crear.</summary>
    public List<SquashProyectoDto> SquashDisponibles { get; set; } = new();
    /// <summary>Crear un proyecto sin Squash (nombre manual) es una salida reservada al administrador.</summary>
    public bool PuedeManual => User.IsInRole(Roles.Administrador);
    public string? SquashError { get; set; }
    private bool _squashCargado;

    private async Task CargarSquashAsync()
    {
        if (_squashCargado) return;
        _squashCargado = true;
        var r = await _squash.ListarProyectosAsync();
        SquashProyectos = r.Datos ?? new();
        SquashError = r.Ok ? null : r.Error;
        var vinculados = await _context.Proyectos.Where(p => p.SquashProjectId != null).Select(p => p.SquashProjectId!.Value).ToListAsync();
        SquashDisponibles = SquashProyectos.Where(p => !vinculados.Contains(p.Id)).ToList();
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

        await CargarSquashAsync();
        var errorSquash = await SquashSincronizacion.ValidarVinculoAsync(_context, Proyecto.SquashProjectId, 0, SquashError is null ? SquashProyectos : null);
        if (errorSquash is not null) ModelState.AddModelError($"{nameof(Proyecto)}.{nameof(Proyecto.SquashProjectId)}", errorSquash);

        if (Proyecto.SquashProjectId is int vinculado)
        {
            // El nombre es el de Squash TM; lo que llegue escrito a mano se ignora.
            var nombreSquash = SquashProyectos.FirstOrDefault(p => p.Id == vinculado)?.Nombre;
            if (!string.IsNullOrWhiteSpace(nombreSquash))
            {
                Proyecto.Nombre = nombreSquash.Length <= 200 ? nombreSquash : nombreSquash[..200];
                ModelState.Remove($"{nameof(Proyecto)}.{nameof(Proyecto.Nombre)}");
            }
        }
        else if (!PuedeManual)
        {
            ModelState.AddModelError($"{nameof(Proyecto)}.{nameof(Proyecto.SquashProjectId)}", "Selecciona el proyecto de Squash TM. Crear un proyecto sin Squash solo lo puede hacer un administrador.");
        }
        else if (SquashError is null && SquashProyectos.Any(p => string.Equals(p.Nombre.Trim(), (Proyecto.Nombre ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase)))
        {
            ModelState.AddModelError($"{nameof(Proyecto)}.{nameof(Proyecto.Nombre)}", "Ese proyecto ya existe en Squash TM: selecciónalo de la lista en lugar de crearlo a mano.");
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
        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateException) when (Proyecto.SquashProjectId is not null)
        {
            // Otro usuario vinculó ese proyecto de Squash un instante antes (índice único).
            ModelState.AddModelError($"{nameof(Proyecto)}.{nameof(Proyecto.SquashProjectId)}", "Ese proyecto de Squash TM acaba de ser vinculado a otro proyecto del portafolio.");
            _context.ChangeTracker.Clear();
            await CargarCatalogosAsync();
            return Page();
        }
        var avisos = new List<string>();
        foreach (var a in Proyecto.Asignaciones)
        {
            var resumen = await _capacidad.EvaluarAsync(a.RecursoQAId, a.FechaInicio, a.FechaFin);
            if (resumen is { Sobreasignado: true })
            {
                avisos.Add($"{resumen.Recurso.NombreCompleto} queda sobreasignado/a en {resumen.DiasSobreasignados} día(s) laborable(s)");
            }
        }

        // La iniciativa ya está guardada: si Bugzilla falla se conserva y se reintenta desde el detalle.
        var bugzilla = await PrepararBugzillaAsync();

        TempData["Mensaje"] = "Proyecto registrado correctamente."
            + (avisos.Count > 0 ? " Atención: " + string.Join("; ", avisos) + "." : "") + bugzilla;
        return RedirectToPage("Index");
    }

    private async Task<string> PrepararBugzillaAsync()
    {
        if (Proyecto.SquashProjectId is null) return string.Empty;
        try
        {
            var r = await _bugzilla.PrepararAsync(Proyecto.Id, false, HttpContext.RequestAborted);
            return r.Estado == EstadoBugzilla.Configurado ? " Producto de Bugzilla sincronizado." : " Bugzilla quedó " + (r.Estado == EstadoBugzilla.Error ? "con error" : "pendiente") + ": " + r.Mensaje + " Reintenta desde el detalle.";
        }
        catch (Exception) { return " Bugzilla quedó pendiente; reintenta desde el detalle."; }
    }

    private async Task CargarCatalogosAsync()
    {
        Instituciones = await _context.Instituciones.Where(i => i.Activo).OrderBy(i => i.Nombre).ToListAsync();
        Sistemas = await _context.Sistemas.Where(s => s.Activo).OrderBy(s => s.Nombre).ToListAsync();
        Recursos = await _context.RecursosQA.Where(r => r.Activo && r.UsuarioId != null).OrderBy(r => r.Nombres).ToListAsync();
        await CargarSquashAsync();
    }
}

public class AsignacionInicial
{
    public int RecursoQAId { get; set; }
    public string Rol { get; set; } = "Analista QA";
    public decimal Porcentaje { get; set; } = 100;
}

