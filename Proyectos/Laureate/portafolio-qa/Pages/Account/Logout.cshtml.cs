using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Account;

public class LogoutModel : PageModel
{
    private readonly SignInManager<UsuarioApp> _signIn;

    public LogoutModel(SignInManager<UsuarioApp> signIn)
    {
        _signIn = signIn;
    }

    // El cierre de sesión es solo por POST (con token antiforgery) para que no se dispare desde un enlace externo.
    public IActionResult OnGet() => RedirectToPage("/Portafolio/Index");

    public async Task<IActionResult> OnPostAsync()
    {
        await _signIn.SignOutAsync();
        return RedirectToPage("/Account/Login");
    }
}
