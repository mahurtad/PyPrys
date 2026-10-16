using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

/// <summary>Los valores numéricos se persisten en BD: no reordenar. "Vencido" no se guarda: se deriva de la fecha límite.</summary>
public enum EstadoAcuerdo
{
    Pendiente = 0,
    EnCurso = 1,
    Cumplido = 2,
    Cancelado = 3
}

public class Acuerdo
{
    public int Id { get; set; }

    [Required]
    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    /// <summary>Correlativo dentro del proyecto (se muestra como AC-001).</summary>
    public int Numero { get; set; }

    [Required(ErrorMessage = "El título es obligatorio.")]
    [StringLength(200)]
    public string Titulo { get; set; } = string.Empty;

    /// <summary>Responsable QA (seleccionable entre los asignados al proyecto). Excluyente con <see cref="ResponsableNombre"/>.</summary>
    public int? RecursoQAId { get; set; }
    public RecursoQA? RecursoQA { get; set; }

    /// <summary>Responsable que no es de QA (PM, líder, negocio…), en texto libre.</summary>
    [StringLength(150)]
    public string? ResponsableNombre { get; set; }

    [DataType(DataType.Date)]
    public DateTime FechaAcuerdo { get; set; } = DateTime.Today;

    [DataType(DataType.Date)]
    public DateTime FechaLimite { get; set; } = DateTime.Today;

    public EstadoAcuerdo Estado { get; set; } = EstadoAcuerdo.Pendiente;

    public List<AcuerdoActividad> Actividades { get; set; } = new();

    public string Codigo => $"AC-{Numero:000}";
    public string Responsable => RecursoQA?.NombreCompleto ?? ResponsableNombre ?? "—";
    public bool Vencido => (Estado == EstadoAcuerdo.Pendiente || Estado == EstadoAcuerdo.EnCurso) && FechaLimite.Date < DateTime.Today;
}

public class AcuerdoActividad
{
    public int AcuerdoId { get; set; }
    public Acuerdo Acuerdo { get; set; } = null!;

    public int ActividadId { get; set; }
    public Actividad Actividad { get; set; } = null!;
}
