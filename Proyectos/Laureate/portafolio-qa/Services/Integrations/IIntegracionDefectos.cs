namespace PortafolioQA.Services.Integrations;

/// <summary>
/// Defecto normalizado proveniente de una fuente externa (Bugzilla).
/// Contrato previsto para la etapa 5; sin implementación real todavía.
/// </summary>
public class DefectoExterno
{
    public string Id { get; set; } = string.Empty;
    public string Titulo { get; set; } = string.Empty;
    public string Estado { get; set; } = string.Empty;
    public string? Resolucion { get; set; }
    public string? Severidad { get; set; }
    public string? Prioridad { get; set; }
    public string? Responsable { get; set; }
    public string? Url { get; set; }
}

public class ResultadoDefectos
{
    public bool Disponible { get; set; }
    public string? Mensaje { get; set; }
    public List<DefectoExterno> Defectos { get; set; } = new();
    public DateTime? UltimaActualizacion { get; set; }
    public bool Desactualizado { get; set; }
}

/// <summary>
/// Contrato para consultar defectos de un proyecto en una fuente externa.
/// La implementación real (Bugzilla) se agrega en la etapa 5.
/// </summary>
public interface IIntegracionDefectos
{
    Task<ResultadoDefectos> ObtenerDefectosAsync(string identificadorExterno, CancellationToken cancellationToken = default);
}

public class IntegracionDefectosNoConfigurada : IIntegracionDefectos
{
    public Task<ResultadoDefectos> ObtenerDefectosAsync(string identificadorExterno, CancellationToken cancellationToken = default)
    {
        return Task.FromResult(new ResultadoDefectos
        {
            Disponible = false,
            Mensaje = "Integración con Bugzilla no configurada todavía."
        });
    }
}
