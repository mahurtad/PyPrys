using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Data;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Usuarios;

public class EditModel : PageModel
{
    private readonly UserManager<UsuarioApp> _usuarios;
    private readonly PortafolioQAContext _context;

    public EditModel(UserManager<UsuarioApp> usuarios, PortafolioQAContext context)
    {
        _usuarios = usuarios;
        _context = context;
    }

    public class EntradaModel
    {
        [Required(ErrorMessage = "Los nombres son obligatorios.")]
        [StringLength(100)]
        public string Nombres { get; set; } = string.Empty;

        [Required(ErrorMessage = "Los apellidos son obligatorios.")]
        [StringLength(100)]
        public string Apellidos { get; set; } = string.Empty;

        [Required(ErrorMessage = "El rol es obligatorio.")]
        public string Rol { get; set; } = Roles.Lector;

        public bool Activo { get; set; } = true;

        [DataType(DataType.Password)]
        public string? NuevaContrasena { get; set; }
    }

    [BindProperty]
    public EntradaModel Entrada { get; set; } = new();

    public string Id { get; set; } = "";
    public string Correo { get; set; } = "";
    public bool EsUnoMismo { get; set; }
    public string[] RolesDisponibles => Roles.Todos;

    public async Task<IActionResult> OnGetAsync(string id)
    {
        var usuario = await _usuarios.FindByIdAsync(id);
        if (usuario is null) return NotFound();

        await CargarAsync(usuario);
        Entrada = new EntradaModel
        {
            Nombres = usuario.Nombres,
            Apellidos = usuario.Apellidos,
            Rol = (await _usuarios.GetRolesAsync(usuario)).FirstOrDefault() ?? Roles.Lector,
            Activo = usuario.Activo
        };
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(string id)
    {
        var usuario = await _usuarios.FindByIdAsync(id);
        if (usuario is null) return NotFound();
        await CargarAsync(usuario);

        if (!Roles.Todos.Contains(Entrada.Rol)) ModelState.AddModelError("Entrada.Rol", "Rol no válido.");

        var rolActual = (await _usuarios.GetRolesAsync(usuario)).FirstOrDefault();
        if (EsUnoMismo && (!Entrada.Activo || Entrada.Rol != Roles.Administrador))
        {
            ModelState.AddModelError(string.Empty, "No puedes desactivarte ni quitarte el rol de administrador a ti mismo; pídelo a otro administrador.");
        }

        // Siempre debe quedar al menos un administrador activo.
        var dejaDeSerAdminActivo = rolActual == Roles.Administrador && usuario.Activo
                                   && (Entrada.Rol != Roles.Administrador || !Entrada.Activo);
        if (dejaDeSerAdminActivo)
        {
            var admins = await _usuarios.GetUsersInRoleAsync(Roles.Administrador);
            if (admins.Count(a => a.Activo && a.Id != usuario.Id) == 0)
            {
                ModelState.AddModelError(string.Empty, "Debe quedar al menos un administrador activo.");
            }
        }

        if (!ModelState.IsValid) return Page();

        usuario.Nombres = Entrada.Nombres.Trim();
        usuario.Apellidos = Entrada.Apellidos.Trim();
        var cambioActivo = usuario.Activo != Entrada.Activo;
        usuario.Activo = Entrada.Activo;

        var actualizar = await _usuarios.UpdateAsync(usuario);
        if (!actualizar.Succeeded) return ConErrores(actualizar);

        // Si el usuario es un recurso QA, mantiene sincronizados nombre y correo.
        var recurso = await _context.RecursosQA.FirstOrDefaultAsync(r => r.UsuarioId == usuario.Id);
        if (recurso is not null)
        {
            recurso.Nombres = usuario.Nombres;
            recurso.Apellidos = usuario.Apellidos;
            await _context.SaveChangesAsync();
        }

        if (rolActual != Entrada.Rol)
        {
            if (rolActual != null) await _usuarios.RemoveFromRoleAsync(usuario, rolActual);
            await _usuarios.AddToRoleAsync(usuario, Entrada.Rol);
        }

        if (!string.IsNullOrEmpty(Entrada.NuevaContrasena))
        {
            var token = await _usuarios.GeneratePasswordResetTokenAsync(usuario);
            var reset = await _usuarios.ResetPasswordAsync(usuario, token, Entrada.NuevaContrasena);
            if (!reset.Succeeded) return ConErrores(reset);
        }

        // Invalida las sesiones abiertas para que el cambio de rol/estado aplique de inmediato.
        if (cambioActivo || rolActual != Entrada.Rol || !string.IsNullOrEmpty(Entrada.NuevaContrasena))
        {
            await _usuarios.UpdateSecurityStampAsync(usuario);
        }

        TempData["Mensaje"] = $"Usuario {usuario.Email} actualizado.";
        return RedirectToPage("Index");
    }

    private IActionResult ConErrores(IdentityResult resultado)
    {
        foreach (var e in resultado.Errors) ModelState.AddModelError(string.Empty, e.Description);
        return Page();
    }

    private Task CargarAsync(UsuarioApp usuario)
    {
        Id = usuario.Id;
        Correo = usuario.Email ?? usuario.UserName ?? "";
        EsUnoMismo = _usuarios.GetUserId(User) == usuario.Id;
        return Task.CompletedTask;
    }
}
