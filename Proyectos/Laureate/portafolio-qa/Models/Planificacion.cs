using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public class Feriado
{
    public int Id { get; set; }

    [Required(ErrorMessage = "La fecha es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime Fecha { get; set; } = DateTime.Today;

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;
}

public class Ausencia
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El recurso QA es obligatorio.")]
    public int RecursoQAId { get; set; }
    public RecursoQA? RecursoQA { get; set; }

    [Required(ErrorMessage = "La fecha de inicio es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaInicio { get; set; } = DateTime.Today;

    [Required(ErrorMessage = "La fecha de fin es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaFin { get; set; } = DateTime.Today;

    [Display(Name = "Tipo")]
    public TipoAusencia Tipo { get; set; } = TipoAusencia.Vacaciones;

    [StringLength(150)]
    public string? Motivo { get; set; }
}

public class Hito
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El proyecto es obligatorio.")]
    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    [Required(ErrorMessage = "La fecha es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime Fecha { get; set; } = DateTime.Today;

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150)]
    public string Nombre { get; set; } = string.Empty;
}
