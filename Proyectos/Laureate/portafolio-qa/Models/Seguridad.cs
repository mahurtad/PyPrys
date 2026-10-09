using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.AspNetCore.Identity;

namespace PortafolioQA.Models;

/// <summary>Usuario de acceso a la aplicación. Es independiente de los recursos QA (un QA puede no tener cuenta).</summary>
public class UsuarioApp : IdentityUser
{
    [StringLength(100)]
    public string Nombres { get; set; } = string.Empty;

    [StringLength(100)]
    public string Apellidos { get; set; } = string.Empty;

    [NotMapped]
    public string NombreCompleto => $"{Nombres} {Apellidos}".Trim();

    /// <summary>Un usuario desactivado conserva su historial pero no puede iniciar sesión.</summary>
    public bool Activo { get; set; } = true;
}

public static class Roles
{
    public const string Administrador = "Administrador";
    public const string Editor = "Editor";
    public const string Lector = "Lector";

    public static readonly string[] Todos = { Administrador, Editor, Lector };

    public const string AdminOEditor = Administrador + "," + Editor;
}

public static class Politicas
{
    public const string PuedeEditar = "PuedeEditar";
    public const string SoloAdmin = "SoloAdmin";
}
