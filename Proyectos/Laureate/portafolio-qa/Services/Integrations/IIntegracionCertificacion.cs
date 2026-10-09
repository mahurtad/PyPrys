namespace PortafolioQA.Services.Integrations;

/// <summary>
/// Resultado normalizado de una consulta de certificación a una fuente externa
/// (por ejemplo Squash TM). Etapa 4 implementará un cliente real; por ahora
/// solo se define el contrato para que el resto de la aplicación pueda
/// consumirlo sin cambios cuando exista una implementación real.
/// </summary>
public class ResultadoCertificacion
{
    public bool Disponible { get; set; }
    public string? Mensaje { get; set; }
    public int? ElementosPlanificados { get; set; }
    public int? ElementosEjecutados { get; set; }
    public int? Aprobados { get; set; }
    public int? Fallidos { get; set; }
    public int? Bloqueados { get; set; }
    public int? Pendientes { get; set; }
    public DateTime? UltimaActualizacion { get; set; }
    public bool Desactualizado { get; set; }
}

/// <summary>
/// Contrato para consultar el avance de certificación de un proyecto en una
/// herramienta externa de gestión de pruebas. La implementación real
/// (Squash TM) se agrega en la etapa 4, sin modificar quien la consume.
/// </summary>
public interface IIntegracionCertificacion
{
    Task<ResultadoCertificacion> ObtenerAvanceAsync(string identificadorExterno, CancellationToken cancellationToken = default);
}

/// <summary>
/// Implementación por defecto mientras no exista una conexión real
/// configurada. Nunca debe reportar cero casos: informa que no hay datos.
/// </summary>
public class IntegracionCertificacionNoConfigurada : IIntegracionCertificacion
{
    public Task<ResultadoCertificacion> ObtenerAvanceAsync(string identificadorExterno, CancellationToken cancellationToken = default)
    {
        return Task.FromResult(new ResultadoCertificacion
        {
            Disponible = false,
            Mensaje = "Integración con Squash TM no configurada todavía."
        });
    }
}
