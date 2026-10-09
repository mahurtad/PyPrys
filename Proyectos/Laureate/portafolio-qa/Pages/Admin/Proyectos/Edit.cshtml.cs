using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;

namespace PortafolioQA.Pages.Admin.Proyectos;

public class EditModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public EditModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public Proyecto Proyecto { get; set; } = new();

    [BindProperty]
    public List<int> SistemasSeleccionados { get; set; } = new();

    public List<Institucion> Instituciones { get; set; } = new();
    public List<Sistema> Sistemas { get; set; } = new();
    public List<Actividad> Actividades { get; set; } = new();
    /// <summary>Recursos QA distintos asignados al proyecto: los únicos que pueden responsabilizarse de una actividad.</summary>
    public List<RecursoQA> RecursosProyecto { get; set; } = new();
    /// <summary>QA que se muestran en el formulario: las asignaciones vigentes o, en un proyecto cerrado sin ellas, el equipo del cierre.</summary>
    public List<string> QAAsignados { get; set; } = new();
    public List<ActividadAuditoria> Auditoria { get; set; } = new();

    private bool EntradaInvalida(params string[] campos) => ModelState.Any(e =>
        campos.Any(c => e.Key.Equals(c, StringComparison.OrdinalIgnoreCase) || e.Key.StartsWith(c + "[", StringComparison.OrdinalIgnoreCase))
        && e.Value?.Errors.Count > 0);

    private string Usuario => User.Identity?.Name ?? CierreProyecto.UsuarioSistema;

    /// <summary>En un proyecto cerrado las actividades son historial: solo un administrador puede corregirlas.</summary>
    public bool ActividadesBloqueadas => Proyecto.Estado == EstadoProyecto.Cerrado && !User.IsInRole(Roles.Administrador);

    private bool Bloqueado(Proyecto proyecto)
    {
        if (proyecto.Estado != EstadoProyecto.Cerrado || User.IsInRole(Roles.Administrador)) return false;
        TempData["Error"] = "El proyecto está cerrado: las actividades son historial y solo un administrador puede corregirlas.";
        return true;
    }

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var proyecto = await _context.Proyectos
            .Include(p => p.Sistemas)
            .FirstOrDefaultAsync(p => p.Id == id);

        if (proyecto is null)
        {
            return NotFound();
        }

        Proyecto = proyecto;
        SistemasSeleccionados = proyecto.Sistemas.Select(s => s.SistemaId).ToList();
        await CargarCatalogosAsync(proyecto.InstitucionId);
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (Proyecto.FechaInicio > Proyecto.FechaFin)
        {
            ModelState.AddModelError($"{nameof(Proyecto)}.{nameof(Proyecto.FechaFin)}", "La fecha de fin debe ser igual o posterior a la fecha de inicio.");
        }

        if (!ModelState.IsValid)
        {
            await CargarCatalogosAsync(Proyecto.InstitucionId);
            return Page();
        }

        var proyecto = await _context.Proyectos
            .Include(p => p.Sistemas)
            .FirstOrDefaultAsync(p => p.Id == Proyecto.Id);

        if (proyecto is null)
        {
            return NotFound();
        }

        proyecto.Nombre = Proyecto.Nombre;
        proyecto.Descripcion = Proyecto.Descripcion;
        proyecto.InstitucionId = Proyecto.InstitucionId;
        proyecto.Tipo = Proyecto.Tipo;
        proyecto.PM = Proyecto.PM;
        proyecto.Lider = Proyecto.Lider;
        proyecto.FechaInicio = Proyecto.FechaInicio;
        proyecto.FechaFin = Proyecto.FechaFin;
        proyecto.Activo = Proyecto.Activo;

        proyecto.Sistemas.Clear();
        foreach (var sistemaId in SistemasSeleccionados)
        {
            proyecto.Sistemas.Add(new ProyectoSistema { ProyectoId = proyecto.Id, SistemaId = sistemaId });
        }

        await _context.SaveChangesAsync();
        var fuera = await _context.Actividades.CountAsync(a => a.ProyectoId == proyecto.Id && a.Estado != EstadoActividad.Cancelada
            && (a.FechaInicio < proyecto.FechaInicio || a.FechaFin > proyecto.FechaFin));
        // Recalcula el avance (si hay actividades) y alinea el estado con el avance; Cancelado no se toca.
        var aviso = await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);

        TempData["Mensaje"] = "Proyecto actualizado correctamente." + aviso
            + (fuera > 0 ? $" Atención: {fuera} actividad(es) quedan fuera del nuevo periodo; ajusta sus fechas." : "");
        return RedirectToPage("Index");
    }

    public async Task<IActionResult> OnPostAgregarActividadAsync(int id, string? actNombre, DateTime? actInicio, DateTime? actFin, decimal? actAvance, int? actRecursoId)
    {
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (proyecto is null) return NotFound();
        if (Bloqueado(proyecto)) return RedirectToPage(null, null, new { id }, "actividades");

        var nombre = actNombre?.Trim();
        var inicio = (actInicio ?? proyecto.FechaInicio).Date;
        var fin = (actFin ?? proyecto.FechaFin).Date;
        var (recursoOk, recursoElegido) = await RecursoValidoAsync(id, actRecursoId);
        if (string.IsNullOrWhiteSpace(nombre) || nombre.Length > 200)
        {
            TempData["Error"] = "El nombre de la actividad es obligatorio (máximo 200 caracteres).";
        }
        else if (EntradaInvalida("actInicio", "actFin"))
        {
            TempData["Error"] = "Las fechas de la actividad no son válidas.";
        }
        else if (EntradaInvalida("actAvance"))
        {
            TempData["Error"] = "El avance debe ser un número válido.";
        }
        else if (EntradaInvalida("actRecursoId"))
        {
            TempData["Error"] = "El recurso QA seleccionado no es válido.";
        }
        else if (inicio > fin)
        {
            TempData["Error"] = "La fecha de fin de la actividad debe ser igual o posterior a la de inicio.";
        }
        else if (actAvance is < 0 or > 100)
        {
            TempData["Error"] = "El avance de la actividad debe estar entre 0 y 100.";
        }
        else if (!recursoOk)
        {
            TempData["Error"] = recursoElegido is null
                ? "Indica el recurso QA responsable de la actividad."
                : "El recurso QA seleccionado no está asignado a este proyecto.";
        }
        else
        {
            var avance = actAvance ?? 0;
            var estadoNuevo = CierreProyecto.EstadoActividadPorAvance(avance);
            var nueva = new Actividad { ProyectoId = id, Nombre = nombre, Orden = 4, FechaInicio = inicio, FechaFin = fin, Estado = estadoNuevo, Avance = avance, RecursoQAId = recursoElegido };
            _context.Actividades.Add(nueva);
            var extendido = CierreProyecto.AjustarFinProyecto(proyecto, fin);
            CierreProyecto.Auditar(_context, proyecto, nueva, nombre, Usuario, "Creada", $"Avance {avance:0.##} %, {estadoNuevo.Texto()}");
            await _context.SaveChangesAsync();
            var aviso = await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
            TempData["Mensaje"] = "Actividad registrada."
                + (extendido ? $" La fecha fin del proyecto se actualizó al {proyecto.FechaFin:dd/MM/yyyy}." : "") + aviso;
        }
        return RedirectToPage(null, null, new { id }, "actividades");
    }

    public async Task<IActionResult> OnPostCancelarActividadAsync(int id, int actividadId)
    {
        var actividad = await _context.Actividades.FirstOrDefaultAsync(a => a.Id == actividadId && a.ProyectoId == id);
        if (actividad is null) return NotFound();
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (proyecto is not null && Bloqueado(proyecto)) return RedirectToPage(null, null, new { id }, "actividades");
        var estadoAnterior = actividad.Estado;
        var estado = actividad.Estado == EstadoActividad.Cancelada
            ? CierreProyecto.EstadoActividadPorAvance(actividad.Avance) : EstadoActividad.Cancelada;
        actividad.Estado = estado;
        if (proyecto is not null) CierreProyecto.Auditar(_context, proyecto, actividad, actividad.Nombre, Usuario, "Estado", $"{estadoAnterior.Texto()} → {estado.Texto()}");
        await _context.SaveChangesAsync();
        if (proyecto is not null)
        {
            var aviso = await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
            if (aviso.Length > 0) TempData["Mensaje"] = aviso.Trim();
        }
        return RedirectToPage(null, null, new { id }, "actividades");
    }

    public async Task<IActionResult> OnPostRecursoActividadAsync(int id, int actividadId, int? recursoId)
    {
        var actividad = await _context.Actividades.FirstOrDefaultAsync(a => a.Id == actividadId && a.ProyectoId == id);
        if (actividad is null) return NotFound();
        var proyectoRec = await _context.Proyectos.FindAsync(id);
        if (proyectoRec is not null && Bloqueado(proyectoRec)) return RedirectToPage(null, null, new { id }, "actividades");
        var (ok, elegido) = await RecursoValidoAsync(id, recursoId, actividad.EsPredeterminada);
        if (EntradaInvalida("recursoId")) ok = false;
        if (!ok)
        {
            TempData["Error"] = elegido is null
                ? "Indica el recurso QA responsable de la actividad."
                : "El recurso QA seleccionado no está asignado a este proyecto.";
        }
        else
        {
            var anterior = actividad.RecursoQAId is null ? "sin asignar"
                : (await _context.RecursosQA.FindAsync(actividad.RecursoQAId))?.NombreCompleto ?? "—";
            var nuevo = elegido is null ? "sin asignar" : (await _context.RecursosQA.FindAsync(elegido))?.NombreCompleto ?? "—";
            actividad.RecursoQAId = elegido;
            if (proyectoRec is not null) CierreProyecto.Auditar(_context, proyectoRec, actividad, actividad.Nombre, Usuario, "Responsable", $"{anterior} → {nuevo}");
            await _context.SaveChangesAsync();
        }
        return RedirectToPage(null, null, new { id }, "actividades");
    }

    public async Task<IActionResult> OnPostAvanceActividadAsync(int id, int actividadId, decimal avance)
    {
        var actividad = await _context.Actividades.FirstOrDefaultAsync(a => a.Id == actividadId && a.ProyectoId == id);
        if (actividad is null) return NotFound();
        var proyectoAv = await _context.Proyectos.FindAsync(id);
        if (proyectoAv is not null && Bloqueado(proyectoAv)) return RedirectToPage(null, null, new { id }, "actividades");
        if (EntradaInvalida("avance") || avance is < 0 or > 100)
        {
            TempData["Error"] = "El avance de la actividad debe estar entre 0 y 100.";
        }
        else
        {
            var avanceAnterior = actividad.Avance;
            actividad.Avance = avance;
            if (actividad.Estado != EstadoActividad.Cancelada) actividad.Estado = CierreProyecto.EstadoActividadPorAvance(avance);
            if (proyectoAv is not null) CierreProyecto.Auditar(_context, proyectoAv, actividad, actividad.Nombre, Usuario, "Avance", $"{avanceAnterior:0.##} % → {avance:0.##} %");
            await _context.SaveChangesAsync();
            var proyecto = proyectoAv;
            if (proyecto is not null)
            {
                var aviso = await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
                if (aviso.Length > 0) TempData["Mensaje"] = aviso.Trim();
            }
        }
        return RedirectToPage(null, null, new { id }, "actividades");
    }

    public async Task<IActionResult> OnPostEliminarActividadAsync(int id, int actividadId)
    {
        var actividad = await _context.Actividades.FirstOrDefaultAsync(a => a.Id == actividadId && a.ProyectoId == id);
        if (actividad is not null)
        {
            if (actividad.EsPredeterminada)
            {
                TempData["Error"] = "Las actividades predeterminadas no se pueden eliminar.";
                return RedirectToPage(null, null, new { id }, "actividades");
            }
            var proyecto = await _context.Proyectos.FindAsync(id);
            if (proyecto is not null && Bloqueado(proyecto)) return RedirectToPage(null, null, new { id }, "actividades");
            if (proyecto is not null) CierreProyecto.Auditar(_context, proyecto, actividad, actividad.Nombre, Usuario, "Eliminada", $"Avance {actividad.Avance:0.##} %, {actividad.Estado.Texto()}");
            _context.Actividades.Remove(actividad);
            await _context.SaveChangesAsync();
            var aviso = proyecto is null ? string.Empty : await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
            TempData["Mensaje"] = "Actividad eliminada." + aviso;
        }
        return RedirectToPage(null, null, new { id }, "actividades");
    }

    public async Task<IActionResult> OnPostCancelarProyectoAsync(int id)
    {
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (proyecto is null) return NotFound();
        if (proyecto.Estado == EstadoProyecto.Cancelado)
        {
            proyecto.Estado = EstadoProyecto.Planificado;
            if (proyecto.AvancePendienteConfirmar) proyecto.Estado = EstadoProyecto.EnCertificacion;
            await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
            TempData["Mensaje"] = "Proyecto reactivado.";
        }
        else
        {
            if (proyecto.Estado == EstadoProyecto.Cerrado)
                await CierreProyecto.ReabrirAsync(_context, proyecto, Usuario);
            proyecto.Estado = EstadoProyecto.Cancelado;
            TempData["Mensaje"] = "Proyecto cancelado.";
        }
        CierreProyecto.Auditar(_context, proyecto, null, proyecto.Nombre, Usuario, "Estado del proyecto", TempData["Mensaje"]?.ToString());
        await _context.SaveChangesAsync();
        return RedirectToPage(new { id });
    }

    // Confirmación explícita y completa: guardar una sola actividad no libera el avance heredado.
    public async Task<IActionResult> OnPostConfirmarAvanceAsync(int id, List<ConfirmacionActividad> confirmaciones)
    {
        var proyecto = await _context.Proyectos.Include(p => p.Actividades).FirstOrDefaultAsync(p => p.Id == id);
        if (proyecto is null) return NotFound();
        if (!proyecto.AvancePendienteConfirmar) return BadRequest();
        if (EntradaInvalida("confirmaciones") || confirmaciones.Count != proyecto.Actividades.Count
            || confirmaciones.Select(c => c.Id).Distinct().Count() != confirmaciones.Count
            || confirmaciones.Any(c => c.Avance is null or < 0 or > 100 || !proyecto.Actividades.Any(a => a.Id == c.Id)))
        {
            TempData["Error"] = "Ingresa un avance válido para cada actividad antes de confirmar.";
            return RedirectToPage(new { id });
        }
        using var transaction = await _context.Database.BeginTransactionAsync();
        foreach (var c in confirmaciones)
        {
            var actividad = proyecto.Actividades.Single(a => a.Id == c.Id);
            actividad.Avance = c.Avance!.Value;
            if (actividad.Estado != EstadoActividad.Cancelada) actividad.Estado = CierreProyecto.EstadoActividadPorAvance(actividad.Avance);
            CierreProyecto.Auditar(_context, proyecto, actividad, actividad.Nombre, Usuario, "Confirmación", $"Avance {actividad.Avance:0.##} %; heredado {proyecto.AvanceLegado:0.##} %");
        }
        proyecto.AvancePendienteConfirmar = false;
        await _context.SaveChangesAsync();
        var aviso = await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
        await transaction.CommitAsync();
        TempData["Mensaje"] = "Avance confirmado." + aviso;
        return RedirectToPage(new { id });
    }

    public async Task<IActionResult> OnPostEditarActividadAsync(int id, int actividadId, string? nombre, DateTime inicio, DateTime fin)
    {
        var actividad = await _context.Actividades.FirstOrDefaultAsync(a => a.Id == actividadId && a.ProyectoId == id);
        var proyecto = await _context.Proyectos.FindAsync(id);
        if (actividad is null || proyecto is null) return NotFound();
        if (Bloqueado(proyecto)) return RedirectToPage(null, null, new { id }, "actividades");
        var errorEdicion =
            actividad.EsPredeterminada && nombre != actividad.Nombre ? "Las actividades predeterminadas no se pueden renombrar."
            : string.IsNullOrWhiteSpace(nombre) || nombre.Trim().Length > 200 ? "El nombre es obligatorio (máximo 200 caracteres)."
            : EntradaInvalida("inicio", "fin") || inicio == default || fin == default ? "Las fechas no son válidas."
            : inicio > fin ? "La fecha de fin debe ser igual o posterior a la de inicio."
            : null;
        if (errorEdicion is not null)
        {
            TempData["Error"] = $"{actividad.Nombre}: {errorEdicion}";
            return RedirectToPage(null, null, new { id }, "actividades");
        }
        actividad.Nombre = nombre.Trim();
        actividad.FechaInicio = inicio.Date;
        actividad.FechaFin = fin.Date;
        CierreProyecto.AjustarFinProyecto(proyecto, fin);
        CierreProyecto.Auditar(_context, proyecto, actividad, actividad.Nombre, Usuario, "Edición", $"{inicio:dd/MM/yyyy} – {fin:dd/MM/yyyy}");
        await _context.SaveChangesAsync();
        TempData["Mensaje"] = "Actividad actualizada." + await CierreProyecto.RecalcularAvanceAsync(_context, proyecto, Usuario);
        return RedirectToPage(null, null, new { id }, "actividades");
    }

    /// <summary>
    /// El responsable debe estar asignado al proyecto. Con un solo recurso se usa ese si no se indica otro; con varios
    /// es obligatorio elegir; sin recursos asignados la actividad queda sin responsable.
    /// </summary>
    private async Task<(bool Ok, int? Elegido)> RecursoValidoAsync(int proyectoId, int? solicitado, bool predeterminada = false)
    {
        var asignados = await _context.Asignaciones.Where(a => a.ProyectoId == proyectoId && !a.NoEjecutada)
            .Select(a => a.RecursoQAId).Distinct().ToListAsync();
        if (solicitado is null && predeterminada) return (true, null);
        if (asignados.Count == 0) return (solicitado is null, solicitado);
        if (solicitado is null) return asignados.Count == 1 ? (true, asignados[0]) : (false, null);
        return (asignados.Contains(solicitado.Value), solicitado);
    }

    private async Task CargarCatalogosAsync(int institucionActualId)
    {
        // También al devolver un formulario inválido: los valores derivados vienen de la BD.
        var derivados = await _context.Proyectos.AsNoTracking().Where(p => p.Id == Proyecto.Id)
            .Select(p => new { p.Estado, p.Avance, p.AvancePendienteConfirmar, p.AvanceLegado }).FirstOrDefaultAsync();
        if (derivados is not null)
        {
            Proyecto.Estado = derivados.Estado;
            Proyecto.Avance = derivados.Avance;
            Proyecto.AvancePendienteConfirmar = derivados.AvancePendienteConfirmar;
            Proyecto.AvanceLegado = derivados.AvanceLegado;
        }
        RecursosProyecto = await _context.Asignaciones.Where(a => a.ProyectoId == Proyecto.Id && !a.NoEjecutada)
            .Select(a => a.RecursoQA!).Distinct().OrderBy(r => r.Nombres).ThenBy(r => r.Apellidos).ToListAsync();
        QAAsignados = RecursosProyecto.Select(r => r.NombreCompleto).ToList();
        if (QAAsignados.Count == 0)
        {
            var cierre = await _context.Cierres.AsNoTracking().Where(c => c.ProyectoId == Proyecto.Id && c.Vigente)
                .OrderByDescending(c => c.Fecha).Select(c => c.Participantes.Where(p => p.Ejecutada).Select(p => p.NombreRecurso).Distinct().ToList()).FirstOrDefaultAsync();
            if (cierre is not null) QAAsignados = cierre.OrderBy(n => n).ToList();
        }
        Actividades = await _context.Actividades.Include(a => a.RecursoQA).Where(a => a.ProyectoId == Proyecto.Id)
            .OrderBy(a => a.Orden).ThenBy(a => a.FechaInicio).ThenBy(a => a.Id).ToListAsync();
        Auditoria = await _context.ActividadAuditorias.AsNoTracking().Where(a => a.ProyectoId == Proyecto.Id)
            .OrderByDescending(a => a.Fecha).Take(30).ToListAsync();
        Instituciones = await _context.Instituciones
            .Where(i => i.Activo || i.Id == institucionActualId)
            .OrderBy(i => i.Nombre).ToListAsync();
        Sistemas = await _context.Sistemas
            .Where(s => s.Activo || SistemasSeleccionados.Contains(s.Id))
            .OrderBy(s => s.Nombre).ToListAsync();
    }
}

public class ConfirmacionActividad
{
    public int Id { get; set; }
    public decimal? Avance { get; set; }
}
