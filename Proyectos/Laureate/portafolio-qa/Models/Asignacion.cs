using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public class Asignacion
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El proyecto es obligatorio.")]
    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    [Required(ErrorMessage = "El recurso QA es obligatorio.")]
    public int RecursoQAId { get; set; }
    public RecursoQA? RecursoQA { get; set; }

    [Required(ErrorMessage = "La fecha de inicio es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaInicio { get; set; } = DateTime.Today;

    [Required(ErrorMessage = "La fecha de fin es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaFin { get; set; } = DateTime.Today.AddMonths(1);

    [Range(0.01, 100, ErrorMessage = "El porcentaje debe estar entre 0 y 100.")]
    public decimal Porcentaje { get; set; } = 100;

    [Required(ErrorMessage = "El rol es obligatorio.")]
    [StringLength(100)]
    public string Rol { get; set; } = string.Empty;

    /// <summary>Planificada pero el proyecto cerró antes de que empezara: se conserva como historial y no consume capacidad.</summary>
    public bool NoEjecutada { get; set; }

    /// <summary>Fin original cuando el cierre del proyecto acortó la asignación; permite restaurarla si se reabre.</summary>
    [DataType(DataType.Date)]
    public DateTime? FechaFinOriginal { get; set; }
}
