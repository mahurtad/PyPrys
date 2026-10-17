namespace PortafolioQA.Services.Integrations;

/// <summary>
/// Capacidad opcional de una fuente de defectos: comprobar que el producto existe y es accesible. Bugzilla responde una lista vacía
/// (no un error) para un producto inexistente o sin permiso, y sin esta comprobación un nombre mal escrito se vería como "cero defectos".
/// Devuelve null si no se pudo comprobar.
/// </summary>
public interface IVerificaProductoDefectos
{
    Task<bool?> ProductoExisteAsync(string producto, CancellationToken cancellationToken = default);
}
