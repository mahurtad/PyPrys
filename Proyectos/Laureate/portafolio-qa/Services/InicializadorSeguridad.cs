using Microsoft.AspNetCore.Identity;
using PortafolioQA.Models;

namespace PortafolioQA.Services;

/// <summary>
/// Crea los roles y, solo si la aplicación no tiene ningún usuario, el primer administrador
/// con las credenciales de configuración (Seguridad:AdminInicial:Correo / Contrasena).
/// Esas credenciales van en user-secrets o variables de entorno, nunca en el repositorio.
/// </summary>
public static class InicializadorSeguridad
{
    public static async Task EjecutarAsync(IServiceProvider servicios)
    {
        using var scope = servicios.CreateScope();
        var logger = scope.ServiceProvider.GetRequiredService<ILoggerFactory>().CreateLogger("Seguridad");
        var config = scope.ServiceProvider.GetRequiredService<IConfiguration>();

        try
        {
            var roles = scope.ServiceProvider.GetRequiredService<RoleManager<IdentityRole>>();
            foreach (var rol in Roles.Todos)
            {
                if (!await roles.RoleExistsAsync(rol)) await roles.CreateAsync(new IdentityRole(rol));
            }

            var usuarios = scope.ServiceProvider.GetRequiredService<UserManager<UsuarioApp>>();
            if (usuarios.Users.Any()) return;

            var correo = config["Seguridad:AdminInicial:Correo"];
            var contrasena = config["Seguridad:AdminInicial:Contrasena"];
            if (string.IsNullOrWhiteSpace(correo) || string.IsNullOrWhiteSpace(contrasena))
            {
                logger.LogWarning("No hay usuarios y no se configuró Seguridad:AdminInicial; nadie podrá iniciar sesión hasta definirlo.");
                return;
            }

            var admin = new UsuarioApp
            {
                UserName = correo, Email = correo, EmailConfirmed = true,
                Nombres = "Administrador", Apellidos = "inicial", Activo = true
            };
            var crear = await usuarios.CreateAsync(admin, contrasena);
            if (!crear.Succeeded)
            {
                logger.LogError("No se pudo crear el administrador inicial: {Errores}",
                    string.Join("; ", crear.Errors.Select(e => e.Description)));
                return;
            }
            await usuarios.AddToRoleAsync(admin, Roles.Administrador);
            logger.LogInformation("Administrador inicial creado: {Correo}. Cambia su contraseña tras el primer ingreso.", correo);
        }
        catch (Exception ex)
        {
            // Típico si aún no se aplicó la migración de la etapa 3.
            logger.LogError(ex, "No se pudo inicializar la seguridad. ¿Se aplicó 'dotnet ef database update'?");
        }
    }
}
