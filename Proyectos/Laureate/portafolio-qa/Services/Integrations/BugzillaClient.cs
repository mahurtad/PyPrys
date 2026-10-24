using System.Text.Json;

namespace PortafolioQA.Services.Integrations;

public class BugzillaOptions
{
    public string? BaseUrl { get; set; }
    public string? ApiKey { get; set; }

    /// <summary>Ruta de la API REST bajo la dirección: "rest" (con reescritura de Apache) o "rest.cgi" (instalaciones sin ella).</summary>
    public string RutaRest { get; set; } = "rest";

    /// <summary>Base de las llamadas REST, con barra final; null si no hay dirección.</summary>
    public string? UrlRest => string.IsNullOrWhiteSpace(BaseUrl) ? null : BaseUrl.TrimEnd('/') + "/" + RutaRest.Trim('/') + "/";
    public bool Configurado => Uri.TryCreate(BaseUrl, UriKind.Absolute, out var uri)
        && uri.Scheme is "http" or "https" && !string.IsNullOrWhiteSpace(ApiKey);

    /// <summary>Clave de la cuenta de servicio con el grupo editcomponents (user-secrets o variable Bugzilla__AdminApiKey).
    /// Es distinta de ApiKey (solo consulta defectos): crear productos requiere un permiso mayor.</summary>
    public string? AdminApiKey { get; set; }

    /// <summary>Login de esa cuenta; es el responsable por defecto del componente "General" (no se asigna a personas).</summary>
    public string CuentaServicio { get; set; } = "portafolio-productos-api@portafolio.local";

    public bool ProductosConfigurado => Uri.TryCreate(BaseUrl, UriKind.Absolute, out var uri)
        && uri.Scheme is "http" or "https" && !string.IsNullOrWhiteSpace(AdminApiKey);
}

/// <summary>Consulta de defectos por nombre exacto de producto de Bugzilla.</summary>
public class BugzillaClient(HttpClient http) : IIntegracionDefectos, IVerificaProductoDefectos
{
    public async Task<ResultadoDefectos> ObtenerDefectosAsync(string identificadorExterno, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(identificadorExterno))
            return Fallo("Indica el nombre del producto en Bugzilla.");
        try
        {
            var defectos = new List<DefectoExterno>();
            const int pagina = 200;
            for (var offset = 0; ; offset += pagina)
            {
                using var response = await http.GetAsync($"bug?product={Uri.EscapeDataString(identificadorExterno.Trim())}&limit={pagina}&offset={offset}&include_fields=id,summary,status,resolution,severity,priority,assigned_to", cancellationToken);
                if (!response.IsSuccessStatusCode)
                    return Fallo(response.StatusCode is System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden
                        ? "Bugzilla rechazó la credencial o los permisos de acceso."
                        : $"Bugzilla respondió {(int)response.StatusCode}.");
                using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync(cancellationToken));
                if (json.RootElement.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.True)
                    return Fallo("Bugzilla no pudo consultar el producto. Revisa su nombre y los permisos.");
                var bugs = json.RootElement.GetProperty("bugs");
                foreach (var bug in bugs.EnumerateArray())
                {
                    var id = bug.GetProperty("id").GetInt32().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string? Campo(string nombre) => bug.TryGetProperty(nombre, out var v) ? v.GetString() : null;
                    defectos.Add(new DefectoExterno { Id = id, Titulo = Campo("summary") ?? "", Estado = Campo("status") ?? "",
                        Resolucion = Campo("resolution"), Severidad = Campo("severity"), Prioridad = Campo("priority"),
                        Responsable = Campo("assigned_to"), Url = new Uri(http.BaseAddress!, "../show_bug.cgi?id=" + id).ToString() });
                }
                if (bugs.GetArrayLength() < pagina) break;
            }
            return new ResultadoDefectos { Disponible = true, Defectos = defectos, UltimaActualizacion = DateTime.UtcNow };
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) { return Fallo("Bugzilla no respondió a tiempo."); }
        catch (Exception ex) when (ex is HttpRequestException or JsonException or KeyNotFoundException or InvalidOperationException)
        { return Fallo("No se pudo consultar Bugzilla. Revisa la conexión y su API REST."); }
    }

    /// <summary>GET product?names=: lista vacía = no existe o la cuenta no tiene acceso.</summary>
    public async Task<bool?> ProductoExisteAsync(string producto, CancellationToken cancellationToken = default)
    {
        try
        {
            using var response = await http.GetAsync($"product?names={Uri.EscapeDataString(producto.Trim())}&include_fields=id", cancellationToken);
            if (!response.IsSuccessStatusCode) return null;
            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync(cancellationToken));
            return json.RootElement.TryGetProperty("products", out var p) ? p.GetArrayLength() > 0 : null;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) { return null; }
        catch (Exception ex) when (ex is HttpRequestException or JsonException or InvalidOperationException) { return null; }
    }

    private static ResultadoDefectos Fallo(string mensaje) => new() { Disponible = false, Mensaje = mensaje };
}
