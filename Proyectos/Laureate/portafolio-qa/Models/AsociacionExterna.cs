using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

/// <summary>
/// Asocia un proyecto local con un proyecto/producto de una fuente externa
/// (Squash TM, Bugzilla). En la etapa 1 no hay conexión real: solo se
/// registra la asociación para dejar el modelo listo para las etapas 4 y 5.
/// </summary>
public class AsociacionExterna
{
    public int Id { get; set; }

    [Required(ErrorMessage = "El proyecto es obligatorio.")]
    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    public HerramientaIntegracion Herramienta { get; set; } = HerramientaIntegracion.SquashTM;

    [Required(ErrorMessage = "El identificador externo es obligatorio.")]
    [StringLength(150)]
    public string IdentificadorExterno { get; set; } = string.Empty;

    public EstadoSincronizacion Estado { get; set; } = EstadoSincronizacion.PendienteConfigurar;

    public DateTime? UltimaSincronizacion { get; set; }

    [StringLength(500)]
    public string? UltimoError { get; set; }
}
