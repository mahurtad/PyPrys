using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public enum EstadoActividad
{
    Pendiente,
    EnCurso,
    Completada,
    Cancelada
}

public class Actividad
{
    public bool EsPredeterminada { get; set; }
    public int Orden { get; set; }

    [Range(typeof(decimal), "0.01", "999.99", ErrorMessage = "El peso debe estar entre 0.01 y 999.99.")]
    public decimal Peso { get; set; } = 10;

    public int Id { get; set; }

    [Required(ErrorMessage = "El proyecto es obligatorio.")]
    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(200)]
    public string Nombre { get; set; } = string.Empty;

    [Required(ErrorMessage = "La fecha de inicio es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaInicio { get; set; } = DateTime.Today;

    [Required(ErrorMessage = "La fecha de fin es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaFin { get; set; } = DateTime.Today;

    [Range(0, 100, ErrorMessage = "El avance debe estar entre 0 y 100.")]
    public decimal Avance { get; set; } = 0;

    /// <summary>Recurso QA responsable; debe estar asignado al proyecto.</summary>
    public int? RecursoQAId { get; set; }
    public RecursoQA? RecursoQA { get; set; }

    public EstadoActividad Estado { get; set; } = EstadoActividad.Pendiente;
}
