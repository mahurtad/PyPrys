using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Usuarios;

public class CreateModel : PageModel
{
    private readonly UserManager<UsuarioApp> _usuarios;

    public CreateModel(UserManager<UsuarioApp> usuarios)
    {
        _usuarios = usuarios;
    }

    public class EntradaModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Correo inválido.")]
        [StringLength(150)]
        public string Correo { get; set; } = string.Empty;

        [Required(ErrorMessage = "Los nombres son obligatorios.")]
        [StringLength(100)]
        public string Nombres { get; set; } = string.Empty;

        [Required(ErrorMessage = "Los apellidos son obligatorios.")]
        [StringLength(100)]
        public string Apellidos { get; set; } = string.Empty;

        [Required(ErrorMessage = "El rol es obligatorio.")]
        public string Rol { get; set; } = Roles.Lector;

        [Required(ErrorMessage = "La contraseña inicial es obligatoria.")]
        [DataType(DataType.Password)]
        public string Contrasena { get; set; } = string.Empty;
    }

    [BindProperty]
    public EntradaModel Entrada { get; set; } = new();

    public string[] RolesDisponibles => Roles.Todos;

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!Roles.Todos.Contains(Entrada.Rol))
        {
            ModelState.AddModelError("Entrada.Rol", "Rol no válido.");
        }
        if (!ModelState.IsValid) return Page();

        var correo = Entrada.Correo.Trim();
        var usuario = new UsuarioApp
        {
            UserName = correo, Email = correo, EmailConfirmed = true,
            Nombres = Entrada.Nombres.Trim(), Apellidos = Entrada.Apellidos.Trim(), Activo = true
        };

        var crear = await _usuarios.CreateAsync(usuario, Entrada.Contrasena);
        if (!crear.Succeeded)
        {
            foreach (var e in crear.Errors) ModelState.AddModelError(string.Empty, e.Description);
            return Page();
        }

        await _usuarios.AddToRoleAsync(usuario, Entrada.Rol);
        TempData["Mensaje"] = $"Usuario {correo} creado con rol {Entrada.Rol}. Entrégale la contraseña inicial por un canal seguro y pídele que la cambie.";
        return RedirectToPage("Index");
    }
}
