using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.RecursosQA;

public class CreateModel : PageModel
{
    private readonly PortafolioQAContext _context;

    public CreateModel(PortafolioQAContext context)
    {
        _context = context;
    }

    [BindProperty]
    public RecursoQA Recurso { get; set; } = new();

    public List<UsuarioApp> Usuarios { get; set; } = new();

    public async Task OnGetAsync()
    {
        await CargarUsuariosAsync(null);
    }

    public async Task<IActionResult> OnPostAsync()
    {
        // Nombres, apellidos y correo no se capturan: provienen del usuario seleccionado.
        ModelState.Remove($"{nameof(Recurso)}.{nameof(Recurso.Nombres)}");
        ModelState.Remove($"{nameof(Recurso)}.{nameof(Recurso.Apellidos)}");
        ModelState.Remove($"{nameof(Recurso)}.{nameof(Recurso.Correo)}");

        var usuario = string.IsNullOrEmpty(Recurso.UsuarioId) ? null
            : await _context.Users.FirstOrDefaultAsync(u => u.Id == Recurso.UsuarioId && u.Activo);
        if (usuario is null)
        {
            ModelState.AddModelError($"{nameof(Recurso)}.{nameof(Recurso.UsuarioId)}", "Selecciona un usuario registrado y activo.");
        }
        else if (await _context.RecursosQA.AnyAsync(r => r.UsuarioId == usuario.Id || r.Correo == usuario.Email))
        {
            ModelState.AddModelError($"{nameof(Recurso)}.{nameof(Recurso.UsuarioId)}", "Ese usuario ya está registrado como recurso QA.");
        }

        if (!ModelState.IsValid || usuario is null)
        {
            await CargarUsuariosAsync(null);
            return Page();
        }

        Recurso.Nombres = usuario.Nombres;
        Recurso.Apellidos = usuario.Apellidos;
        Recurso.Correo = usuario.Email ?? usuario.UserName ?? string.Empty;
        _context.RecursosQA.Add(Recurso);
        await _context.SaveChangesAsync();

        TempData["Mensaje"] = "Recurso QA registrado correctamente.";
        return RedirectToPage("Index");
    }

    private async Task CargarUsuariosAsync(string? usuarioActualId)
    {
        var vinculados = _context.RecursosQA.Where(r => r.UsuarioId != null).Select(r => r.UsuarioId!);
        Usuarios = await _context.Users.AsNoTracking()
            .Where(u => u.Activo && (!vinculados.Contains(u.Id) || u.Id == usuarioActualId))
            .OrderBy(u => u.Nombres).ThenBy(u => u.Apellidos).ToListAsync();
    }
}
