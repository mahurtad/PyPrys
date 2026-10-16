using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.RecursosQA;

public class EditModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public EditModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public RecursoQA Recurso { get; set; } = new();

    public List<UsuarioApp> Usuarios { get; set; } = new();

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var recurso = await _context.RecursosQA.FindAsync(id);
        if (recurso is null)
        {
            return NotFound();
        }

        Recurso = recurso;
        await CargarUsuariosAsync(recurso.UsuarioId);
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        ModelState.Remove($"{nameof(Recurso)}.{nameof(Recurso.Nombres)}");
        ModelState.Remove($"{nameof(Recurso)}.{nameof(Recurso.Apellidos)}");
        ModelState.Remove($"{nameof(Recurso)}.{nameof(Recurso.Correo)}");

        var existente = await _context.RecursosQA.FirstOrDefaultAsync(r => r.Id == Recurso.Id);
        if (existente is null) return NotFound();

        var usuario = string.IsNullOrEmpty(Recurso.UsuarioId) ? null
            : await _context.Users.FirstOrDefaultAsync(u => u.Id == Recurso.UsuarioId && (u.Activo || u.Id == existente.UsuarioId));
        if (usuario is null)
        {
            ModelState.AddModelError($"{nameof(Recurso)}.{nameof(Recurso.UsuarioId)}", "Selecciona un usuario registrado y activo.");
        }
        else if (await _context.RecursosQA.AnyAsync(r => r.Id != Recurso.Id && (r.UsuarioId == usuario.Id || r.Correo == usuario.Email)))
        {
            ModelState.AddModelError($"{nameof(Recurso)}.{nameof(Recurso.UsuarioId)}", "Ese usuario ya está registrado como otro recurso QA.");
        }

        var loginSquash = string.IsNullOrWhiteSpace(Recurso.SquashLogin) ? usuario?.Email : Recurso.SquashLogin.Trim();
        if (!string.IsNullOrWhiteSpace(loginSquash) && await _context.RecursosQA.AnyAsync(r => r.Id != Recurso.Id && r.SquashLogin == loginSquash))
        {
            ModelState.AddModelError($"{nameof(Recurso)}.{nameof(Recurso.SquashLogin)}", "Ese login de Squash TM ya está asignado a otro recurso.");
        }

        if (!ModelState.IsValid || usuario is null)
        {
            await CargarUsuariosAsync(existente.UsuarioId);
            return Page();
        }

        existente.UsuarioId = usuario.Id;
        existente.Nombres = usuario.Nombres;
        existente.Apellidos = usuario.Apellidos;
        existente.Correo = usuario.Email ?? usuario.UserName ?? string.Empty;
        existente.SquashLogin = loginSquash;
        existente.Equipo = Recurso.Equipo;
        existente.CapacidadDisponible = Recurso.CapacidadDisponible;
        existente.Activo = Recurso.Activo;
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Recurso QA actualizado correctamente.";
        return RedirectToPage("Index");
    }

    private async Task CargarUsuariosAsync(string? usuarioActualId)
    {
        var vinculados = _context.RecursosQA.Where(r => r.UsuarioId != null).Select(r => r.UsuarioId!);
        Usuarios = await _context.Users.AsNoTracking()
            .Where(u => (u.Activo && !vinculados.Contains(u.Id)) || u.Id == usuarioActualId)
            .OrderBy(u => u.Nombres).ThenBy(u => u.Apellidos).ToListAsync();
    }
}
