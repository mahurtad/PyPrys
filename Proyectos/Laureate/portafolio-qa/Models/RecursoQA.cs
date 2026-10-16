using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public class RecursoQA
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(100)]
    public string Nombres { get; set; } = string.Empty;

    [Required(ErrorMessage = "El apellido es obligatorio.")]
    [StringLength(100)]
    public string Apellidos { get; set; } = string.Empty;

    [Required(ErrorMessage = "El correo es obligatorio.")]
    [EmailAddress(ErrorMessage = "Correo inválido.")]
    [StringLength(150)]
    public string Correo { get; set; } = string.Empty;

    public EquipoQA Equipo { get; set; } = EquipoQA.Change;

    [Range(0, 100, ErrorMessage = "La capacidad debe estar entre 0 y 100.")]
    public decimal CapacidadDisponible { get; set; } = 100;

    public bool Activo { get; set; } = true;

    /// <summary>Login de esta persona en Squash TM. Vincula los casos que Squash le asigna con este recurso.</summary>
    [Display(Name = "Usuario en Squash TM")]
    [StringLength(100)]
    public string? SquashLogin { get; set; }

    /// <summary>Usuario registrado al que corresponde este recurso. Para asignar a un recurso debe estar vinculado.</summary>
    public string? UsuarioId { get; set; }
    public UsuarioApp? Usuario { get; set; }

    public string NombreCompleto => $"{Nombres} {Apellidos}";

    public List<Asignacion> Asignaciones { get; set; } = new();
}
