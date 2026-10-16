using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Account;

public class CambiarPasswordModel : PageModel
{
    private readonly UserManager<UsuarioApp> _usuarios;
    private readonly SignInManager<UsuarioApp> _signIn;

    public CambiarPasswordModel(UserManager<UsuarioApp> usuarios, SignInManager<UsuarioApp> signIn)
    {
        _usuarios = usuarios;
        _signIn = signIn;
    }

    public class EntradaModel
    {
        [Required(ErrorMessage = "Indica tu contraseña actual.")]
        [DataType(DataType.Password)]
        public string Actual { get; set; } = string.Empty;

        [Required(ErrorMessage = "Indica la nueva contraseña.")]
        [DataType(DataType.Password)]
        public string Nueva { get; set; } = string.Empty;

        [DataType(DataType.Password)]
        [Compare(nameof(Nueva), ErrorMessage = "La confirmación no coincide.")]
        public string Confirmacion { get; set; } = string.Empty;
    }

    [BindProperty]
    public EntradaModel Entrada { get; set; } = new();

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid) return Page();

        var usuario = await _usuarios.GetUserAsync(User);
        if (usuario is null) return RedirectToPage("/Account/Login");

        var resultado = await _usuarios.ChangePasswordAsync(usuario, Entrada.Actual, Entrada.Nueva);
        if (!resultado.Succeeded)
        {
            foreach (var e in resultado.Errors) ModelState.AddModelError(string.Empty, e.Description);
            return Page();
        }

        await _signIn.RefreshSignInAsync(usuario);
        TempData["Mensaje"] = "Contraseña actualizada.";
        return RedirectToPage("/Portafolio/Index");
    }
}
