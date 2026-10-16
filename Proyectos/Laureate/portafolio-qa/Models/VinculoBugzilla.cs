using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

/// <summary>Los valores numéricos se persisten en BD: no reordenar.</summary>
public enum EstadoBugzilla
{
    Pendiente = 0,
    Configurado = 1,
    Error = 2
}

/// <summary>
/// Producto de Bugzilla que corresponde a una iniciativa (1 a 1). Los identificadores externos se guardan junto con el
/// ambiente de origen (URL base normalizada), porque dos instalaciones distintas pueden repetir los mismos IDs numéricos.
/// Una vez guardado el ProductoId la relación no depende del nombre: renombrar la iniciativa no cambia el producto.
/// </summary>
public class VinculoBugzilla
{
    public int Id { get; set; }

    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    /// <summary>Ambiente (URL base normalizada) del Squash TM al que pertenece SquashProjectId.</summary>
    [StringLength(300)]
    public string SquashAmbiente { get; set; } = string.Empty;
    public int SquashProjectId { get; set; }

    /// <summary>Ambiente (URL base normalizada) del Bugzilla al que pertenecen los IDs siguientes.</summary>
    [StringLength(300)]
    public string BugzillaAmbiente { get; set; } = string.Empty;

    public int? ProductoId { get; set; }
    [StringLength(200)]
    public string? ProductoNombre { get; set; }
    public int? ComponenteId { get; set; }
    [StringLength(100)]
    public string? ComponenteNombre { get; set; }
    public int? VersionId { get; set; }
    [StringLength(100)]
    public string? VersionNombre { get; set; }

    public EstadoBugzilla Estado { get; set; } = EstadoBugzilla.Pendiente;

    /// <summary>Última vez que producto, componente y versión quedaron verificados en Bugzilla.</summary>
    public DateTime? UltimaSincronizacion { get; set; }
    public DateTime? UltimoIntento { get; set; }

    [StringLength(500)]
    public string? UltimoError { get; set; }

    /// <summary>Un administrador confirmó que el producto existente con el mismo nombre corresponde a esta iniciativa.</summary>
    public bool AsociacionExistenteConfirmada { get; set; }

    /// <summary>Reserva temporal que impide dos preparaciones simultáneas de la misma iniciativa.</summary>
    public DateTime? BloqueadoHasta { get; set; }

    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;
}
