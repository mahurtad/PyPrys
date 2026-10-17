using System.ComponentModel.DataAnnotations;

namespace PortafolioQA.Models;

// Última lectura de Squash TM (solo lectura). Las filas se identifican por su id de Squash y no se borran:
// lo que desaparece de Squash queda con Vigente = false y deja de contar en el avance.

public class SquashCampana
{
    public int Id { get; set; }

    public int ProyectoId { get; set; }
    public Proyecto? Proyecto { get; set; }

    public int SquashId { get; set; }

    [StringLength(200)]
    public string Nombre { get; set; } = string.Empty;

    [StringLength(30)]
    public string? Estado { get; set; }

    public bool Vigente { get; set; } = true;

    public List<SquashIteracion> Iteraciones { get; set; } = new();
}

/// <summary>Ejecución de pruebas (iteración de una campaña): es el "título" bajo el que se listan los casos.</summary>
public class SquashIteracion
{
    public int Id { get; set; }

    public int SquashCampanaId { get; set; }
    public SquashCampana? SquashCampana { get; set; }

    public int SquashId { get; set; }

    [StringLength(200)]
    public string Nombre { get; set; } = string.Empty;

    [StringLength(30)]
    public string? Estado { get; set; }

    /// <summary>Promedio del porcentaje de sus casos vigentes que cuentan para el avance.</summary>
    public decimal Porcentaje { get; set; }

    public bool Vigente { get; set; } = true;

    public List<SquashCaso> Casos { get; set; } = new();

    public List<SquashModulo> Modulos { get; set; } = new();
}

/// <summary>
/// Módulo = suite de pruebas de Squash dentro de una iteración. Agrupa ítems del plan de pruebas por la relación real
/// que informa Squash (test-suites/{id}/test-plan), no por nombres. Un ítem puede estar en varios módulos o en ninguno.
/// </summary>
public class SquashModulo
{
    public int Id { get; set; }

    public int SquashIteracionId { get; set; }
    public SquashIteracion? SquashIteracion { get; set; }

    /// <summary>Id de la suite en Squash.</summary>
    public int SquashId { get; set; }

    [StringLength(200)]
    public string Nombre { get; set; } = string.Empty;

    public bool Vigente { get; set; } = true;

    public List<SquashCaso> Casos { get; set; } = new();
}

public class SquashCaso
{
    public int Id { get; set; }

    public int SquashIteracionId { get; set; }
    public SquashIteracion? SquashIteracion { get; set; }

    /// <summary>Id del ítem del plan de pruebas (identifica al caso dentro de la iteración).</summary>
    public int SquashItemId { get; set; }

    /// <summary>Id del caso de prueba en la biblioteca de Squash. El mismo caso puede estar en varios ítems/iteraciones.</summary>
    public int? SquashTestCaseId { get; set; }

    [StringLength(300)]
    public string Nombre { get; set; } = string.Empty;

    [StringLength(30)]
    public string EstadoSquash { get; set; } = string.Empty;

    /// <summary>Null = el estado no cuenta para el avance (p. ej. no ejecutable u omitido).</summary>
    public decimal? Porcentaje { get; set; }

    [StringLength(100)]
    public string? AsignadoLogin { get; set; }

    /// <summary>Recurso QA que corresponde al login de Squash (RecursoQA.SquashLogin), si existe.</summary>
    public int? RecursoQAId { get; set; }
    public RecursoQA? RecursoQA { get; set; }

    public DateTime? UltimaEjecucion { get; set; }

    public bool Vigente { get; set; } = true;

    public List<SquashModulo> Modulos { get; set; } = new();
}
