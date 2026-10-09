using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PortafolioQA.Models;

namespace PortafolioQA.Pages.Admin.Usuarios;

public class FilaUsuario
{
    public string Id { get; set; } = "";
    public string Correo { get; set; } = "";
    public string Nombres { get; set; } = "";
    public string Apellidos { get; set; } = "";
    public string Rol { get; set; } = "(sin rol)";
    public bool Activo { get; set; }
    public bool Bloqueado { get; set; }
}

public class IndexModel : PageModel
{
    private readonly UserManager<UsuarioApp> _usuarios;

    public IndexModel(UserManager<UsuarioApp> usuarios)
    {
        _usuarios = usuarios;
    }

    public List<FilaUsuario> Filas { get; set; } = new();

    public async Task OnGetAsync()
    {
        var lista = await _usuarios.Users.AsNoTracking().OrderBy(u => u.Nombres).ThenBy(u => u.Apellidos).ToListAsync();
        foreach (var u in lista)
        {
            var roles = await _usuarios.GetRolesAsync(u);
            Filas.Add(new FilaUsuario
            {
                Id = u.Id,
                Correo = u.Email ?? u.UserName ?? "",
                Nombres = u.Nombres,
                Apellidos = u.Apellidos,
                Rol = roles.FirstOrDefault() ?? "(sin rol)",
                Activo = u.Activo,
                Bloqueado = u.LockoutEnd.HasValue && u.LockoutEnd > DateTimeOffset.UtcNow
            });
        }
    }
}
