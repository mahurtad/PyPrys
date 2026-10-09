using Microsoft.AspNetCore.Mvc.ModelBinding;
using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

public class Proyecto
{
    [BindNever]
    public bool AvancePendienteConfirmar { get; set; }
    [BindNever]
    public decimal? AvanceLegado { get; set; }

    public int Id { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(200)]
    public string Nombre { get; set; } = string.Empty;

    [Display(Name = "Descripción")]
    [StringLength(500)]
    public string? Descripcion { get; set; }

    [Display(Name = "Institución")]
    [Required(ErrorMessage = "La institución es obligatoria.")]
    public int InstitucionId { get; set; }
    public Institucion? Institucion { get; set; }

    public TipoIniciativa Tipo { get; set; } = TipoIniciativa.Proyecto;

    [BindNever]
    public EstadoProyecto Estado { get; set; } = EstadoProyecto.Planificado;

    [StringLength(150)]
    public string? PM { get; set; }

    [Display(Name = "Líder")]
    [StringLength(150)]
    public string? Lider { get; set; }

    [Display(Name = "Fecha de inicio")]
    [Required(ErrorMessage = "La fecha de inicio es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaInicio { get; set; } = DateTime.Today;

    [Display(Name = "Fecha de fin")]
    [Required(ErrorMessage = "La fecha de fin es obligatoria.")]
    [DataType(DataType.Date)]
    public DateTime FechaFin { get; set; } = DateTime.Today.AddMonths(1);

    [Range(0, 100, ErrorMessage = "El avance debe estar entre 0 y 100.")]
    [BindNever]
    public decimal Avance { get; set; } = 0;

    public bool Activo { get; set; } = true;

    public List<ProyectoSistema> Sistemas { get; set; } = new();
    public List<Asignacion> Asignaciones { get; set; } = new();
    public List<Hito> Hitos { get; set; } = new();
    [BindNever]
    public List<Actividad> Actividades { get; set; } = new();
    public List<AsociacionExterna> Integraciones { get; set; } = new();
    [BindNever]
    public List<CierreProyectoRegistro> Cierres { get; set; } = new();
}

public class ProyectoSistema
{
    public int ProyectoId { get; set; }
    public Proyecto Proyecto { get; set; } = null!;

    public int SistemaId { get; set; }
    public Sistema Sistema { get; set; } = null!;
}
