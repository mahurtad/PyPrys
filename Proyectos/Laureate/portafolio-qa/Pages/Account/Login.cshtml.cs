using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Account;

public class LoginModel : PageModel
{
    private readonly SignInManager<UsuarioApp> _signIn;
    private readonly UserManager<UsuarioApp> _usuarios;
    private readonly ILogger<LoginModel> _logger;

    public LoginModel(SignInManager<UsuarioApp> signIn, UserManager<UsuarioApp> usuarios, ILogger<LoginModel> logger)
    {
        _signIn = signIn;
        _usuarios = usuarios;
        _logger = logger;
    }

    public class EntradaModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Correo inválido.")]
        public string Correo { get; set; } = string.Empty;

        [Required(ErrorMessage = "La contraseña es obligatoria.")]
        [DataType(DataType.Password)]
        public string Contrasena { get; set; } = string.Empty;
    }

    [BindProperty]
    public EntradaModel Entrada { get; set; } = new();

    [BindProperty(SupportsGet = true)]
    public string? ReturnUrl { get; set; }

    public IActionResult OnGet()
    {
        if (User.Identity?.IsAuthenticated == true) return LocalRedirect(DestinoSeguro());
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid) return Page();

        // Mensaje único para cualquier fallo: no revela si el correo existe, está desactivado o bloqueado.
        const string mensaje = "Correo o contraseña incorrectos, o la cuenta no está disponible.";

        var usuario = await _usuarios.FindByEmailAsync(Entrada.Correo);
        if (usuario is null || !usuario.Activo)
        {
            ModelState.AddModelError(string.Empty, mensaje);
            return Page();
        }

        var resultado = await _signIn.PasswordSignInAsync(usuario, Entrada.Contrasena, isPersistent: false, lockoutOnFailure: true);
        if (resultado.Succeeded)
        {
            _logger.LogInformation("Inicio de sesión: {Usuario}", usuario.Email);
            return LocalRedirect(DestinoSeguro());
        }

        if (resultado.IsLockedOut) _logger.LogWarning("Cuenta bloqueada por intentos fallidos: {Usuario}", usuario.Email);
        ModelState.AddModelError(string.Empty, mensaje);
        return Page();
    }

    // Solo se redirige a rutas locales: evita el redireccionamiento abierto tras el login.
    private string DestinoSeguro() =>
        !string.IsNullOrEmpty(ReturnUrl) && Url.IsLocalUrl(ReturnUrl) ? ReturnUrl : "/";
}
