using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

/// <summary>Foto inmutable de cómo terminó un proyecto: quién participó y qué actividades tenía al cerrarse.</summary>
public class CierreProyectoRegistro
{
    public int Id { get; set; }

    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    public DateTime Fecha { get; set; }

    [StringLength(256)]
    public string Usuario { get; set; } = string.Empty;

    public decimal AvanceFinal { get; set; }
    public int TotalActividades { get; set; }
    public int ActividadesCompletadas { get; set; }

    /// <summary>Cierre anterior a la trazabilidad: se armó con las asignaciones que sobrevivían; las actividades no conservan responsable.</summary>
    public bool Reconstruido { get; set; }

    /// <summary>False cuando el proyecto se reabrió: el registro queda como historial del cierre previo.</summary>
    public bool Vigente { get; set; } = true;
    public DateTime? FechaReapertura { get; set; }

    [StringLength(256)]
    public string? UsuarioReapertura { get; set; }

    public List<CierreParticipante> Participantes { get; set; } = new();
    public List<CierreActividad> Actividades { get; set; } = new();
}

public class CierreParticipante
{
    public int Id { get; set; }

    public int CierreProyectoRegistroId { get; set; }
    public CierreProyectoRegistro? Cierre { get; set; }

    public int RecursoQAId { get; set; }
    public RecursoQA? RecursoQA { get; set; }

    /// <summary>Nombre congelado al momento del cierre.</summary>
    [StringLength(200)]
    public string NombreRecurso { get; set; } = string.Empty;

    [StringLength(100)]
    public string Rol { get; set; } = string.Empty;

    public decimal Porcentaje { get; set; }

    /// <summary>Periodo real de participación (fin recortado a la fecha de cierre).</summary>
    public DateTime FechaInicio { get; set; }
    public DateTime FechaFin { get; set; }

    /// <summary>False si la asignación estaba planificada pero el proyecto cerró antes de que empezara.</summary>
    public bool Ejecutada { get; set; } = true;
}

public class CierreActividad
{
    public int Id { get; set; }

    public int CierreProyectoRegistroId { get; set; }
    public CierreProyectoRegistro? Cierre { get; set; }

    [StringLength(200)]
    public string Nombre { get; set; } = string.Empty;

    [StringLength(200)]
    public string? Responsable { get; set; }

    [StringLength(30)]
    public string Estado { get; set; } = string.Empty;

    public decimal Avance { get; set; }
    public DateTime FechaInicio { get; set; }
    public DateTime FechaFin { get; set; }
}

/// <summary>Bitácora de cambios sobre actividades (quién, cuándo, qué).</summary>
public class ActividadAuditoria
{
    public int Id { get; set; }

    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    public int? ActividadId { get; set; }

    [StringLength(200)]
    public string Actividad { get; set; } = string.Empty;

    public DateTime Fecha { get; set; }

    [StringLength(256)]
    public string Usuario { get; set; } = string.Empty;

    [StringLength(40)]
    public string Accion { get; set; } = string.Empty;

    [StringLength(500)]
    public string? Detalle { get; set; }

    /// <summary>True si el proyecto estaba Cerrado cuando se hizo el cambio (corrección posterior al cierre).</summary>
    public bool ProyectoCerrado { get; set; }
}
