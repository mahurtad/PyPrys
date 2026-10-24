using System.Net;
using System.Text;
using System.Text.Json;

namespace PortafolioQA.Services.Integrations;

public record ComponenteBugzilla(int Id, string Nombre);
public record VersionBugzilla(int Id, string Nombre);
public record ProductoResumenBugzilla(int Id, string Nombre);
public record ProductoBugzilla(int Id, string Nombre, string? Descripcion, List<ComponenteBugzilla> Componentes, List<VersionBugzilla> Versiones);

public class ResultadoBugzilla<T>
{
    public bool Ok { get; init; }
    public string? Error { get; init; }
    public T? Datos { get; init; }

    public static ResultadoBugzilla<T> Correcto(T datos) => new() { Ok = true, Datos = datos };
    public static ResultadoBugzilla<T> Fallo(string error) => new() { Ok = false, Error = error };
}

/// <summary>
/// Alta de productos en Bugzilla solo mediante su API REST oficial (cuenta de servicio con el grupo editcomponents).
/// Bugzilla 5.2 crea la versión inicial junto con el producto y no ofrece un endpoint para crear versiones sueltas.
/// </summary>
public interface IBugzillaProductos
{
    Task<ResultadoBugzilla<List<ProductoResumenBugzilla>>> ListarProductosAsync(CancellationToken ct = default);
    Task<ResultadoBugzilla<ProductoBugzilla?>> ObtenerProductoAsync(int id, CancellationToken ct = default);
    Task<ResultadoBugzilla<int>> CrearProductoAsync(string nombre, string descripcion, string versionInicial, CancellationToken ct = default);
    Task<ResultadoBugzilla<int>> CrearComponenteAsync(string producto, string nombre, string descripcion, string asignadoPorDefecto, CancellationToken ct = default);
}

public class BugzillaProductosNoConfigurado : IBugzillaProductos
{
    public const string Mensaje = "Bugzilla no está configurado para crear productos (falta la clave de la cuenta de servicio: Sistema → Integración Bugzilla).";
    private static ResultadoBugzilla<T> Fallo<T>() => ResultadoBugzilla<T>.Fallo(Mensaje);
    public Task<ResultadoBugzilla<List<ProductoResumenBugzilla>>> ListarProductosAsync(CancellationToken ct = default) => Task.FromResult(Fallo<List<ProductoResumenBugzilla>>());
    public Task<ResultadoBugzilla<ProductoBugzilla?>> ObtenerProductoAsync(int id, CancellationToken ct = default) => Task.FromResult(Fallo<ProductoBugzilla?>());
    public Task<ResultadoBugzilla<int>> CrearProductoAsync(string nombre, string descripcion, string versionInicial, CancellationToken ct = default) => Task.FromResult(Fallo<int>());
    public Task<ResultadoBugzilla<int>> CrearComponenteAsync(string producto, string nombre, string descripcion, string asignadoPorDefecto, CancellationToken ct = default) => Task.FromResult(Fallo<int>());
}

/// <summary>
/// Este Bugzilla (5.2) ignora la cabecera X-BUGZILLA-API-KEY y trata la llamada como anónima (401 al crear): la clave se envía
/// como parámetro Bugzilla_api_key. Se añade aquí, y el registro HTTP del cliente se silencia en Program.cs para no volcar la URL.
/// </summary>
public class BugzillaApiKeyHandler(Func<string> clave) : DelegatingHandler
{
    public BugzillaApiKeyHandler(string clave) : this(() => clave) { }

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        var uri = request.RequestUri!;
        var separador = string.IsNullOrEmpty(uri.Query) ? "?" : "&";
        request.RequestUri = new Uri(uri.OriginalString + separador + "Bugzilla_api_key=" + Uri.EscapeDataString(clave()), UriKind.RelativeOrAbsolute);
        return base.SendAsync(request, ct);
    }
}

public class BugzillaProductosClient(HttpClient http) : IBugzillaProductos
{
    public async Task<ResultadoBugzilla<List<ProductoResumenBugzilla>>> ListarProductosAsync(CancellationToken ct = default)
    {
        // Se lista lo accesible y se compara por nombre localmente: así "no existe" nunca se confunde con un error de permisos.
        var r = await EnviarAsync(HttpMethod.Get, "product?type=accessible&include_fields=id,name", null, ct);
        if (!r.Ok) return ResultadoBugzilla<List<ProductoResumenBugzilla>>.Fallo(r.Error!);
        try
        {
            var lista = r.Datos!.RootElement.GetProperty("products").EnumerateArray()
                .Select(p => new ProductoResumenBugzilla(p.GetProperty("id").GetInt32(), p.GetProperty("name").GetString() ?? string.Empty)).ToList();
            return ResultadoBugzilla<List<ProductoResumenBugzilla>>.Correcto(lista);
        }
        catch (Exception ex) when (ex is KeyNotFoundException or InvalidOperationException or FormatException)
        { return ResultadoBugzilla<List<ProductoResumenBugzilla>>.Fallo("Bugzilla devolvió una lista de productos que no se pudo interpretar."); }
        finally { r.Datos!.Dispose(); }
    }

    public async Task<ResultadoBugzilla<ProductoBugzilla?>> ObtenerProductoAsync(int id, CancellationToken ct = default)
    {
        var r = await EnviarAsync(HttpMethod.Get, $"product/{id}", null, ct);
        if (!r.Ok) return ResultadoBugzilla<ProductoBugzilla?>.Fallo(r.Error!);
        try
        {
            var productos = r.Datos!.RootElement.GetProperty("products");
            if (productos.GetArrayLength() == 0) return ResultadoBugzilla<ProductoBugzilla?>.Correcto(null);
            var p = productos[0];
            string? Texto(JsonElement e, string n) => e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
            var componentes = p.TryGetProperty("components", out var cs) && cs.ValueKind == JsonValueKind.Array
                ? cs.EnumerateArray().Select(c => new ComponenteBugzilla(c.GetProperty("id").GetInt32(), Texto(c, "name") ?? string.Empty)).ToList() : new();
            var versiones = p.TryGetProperty("versions", out var vs) && vs.ValueKind == JsonValueKind.Array
                ? vs.EnumerateArray().Select(v => new VersionBugzilla(v.GetProperty("id").GetInt32(), Texto(v, "name") ?? string.Empty)).ToList() : new();
            return ResultadoBugzilla<ProductoBugzilla?>.Correcto(new ProductoBugzilla(p.GetProperty("id").GetInt32(), Texto(p, "name") ?? string.Empty, Texto(p, "description"), componentes, versiones));
        }
        catch (Exception ex) when (ex is KeyNotFoundException or InvalidOperationException or FormatException)
        { return ResultadoBugzilla<ProductoBugzilla?>.Fallo("Bugzilla devolvió un producto que no se pudo interpretar."); }
        finally { r.Datos!.Dispose(); }
    }

    public async Task<ResultadoBugzilla<int>> CrearProductoAsync(string nombre, string descripcion, string versionInicial, CancellationToken ct = default) =>
        await CrearAsync("product", new { name = nombre, description = descripcion, version = versionInicial, is_open = true, has_unconfirmed = true }, ct);

    public async Task<ResultadoBugzilla<int>> CrearComponenteAsync(string producto, string nombre, string descripcion, string asignadoPorDefecto, CancellationToken ct = default) =>
        await CrearAsync("component", new { product = producto, name = nombre, description = descripcion, default_assignee = asignadoPorDefecto }, ct);

    private async Task<ResultadoBugzilla<int>> CrearAsync(string ruta, object cuerpo, CancellationToken ct)
    {
        var r = await EnviarAsync(HttpMethod.Post, ruta, cuerpo, ct);
        if (!r.Ok) return ResultadoBugzilla<int>.Fallo(r.Error!);
        using var doc = r.Datos!;
        return doc.RootElement.TryGetProperty("id", out var id) && id.TryGetInt32(out var valor)
            ? ResultadoBugzilla<int>.Correcto(valor)
            : ResultadoBugzilla<int>.Fallo("Bugzilla no devolvió el identificador del elemento creado.");
    }

    /// <summary>Los mensajes de error nunca incluyen la clave ni el cuerpo completo de la respuesta.</summary>
    private async Task<ResultadoBugzilla<JsonDocument>> EnviarAsync(HttpMethod metodo, string ruta, object? cuerpo, CancellationToken ct)
    {
        try
        {
            using var solicitud = new HttpRequestMessage(metodo, ruta);
            if (cuerpo is not null) solicitud.Content = new StringContent(JsonSerializer.Serialize(cuerpo), Encoding.UTF8, "application/json");
            using var respuesta = await http.SendAsync(solicitud, ct);
            var texto = await respuesta.Content.ReadAsStringAsync(ct);
            JsonDocument? doc = null;
            try { doc = JsonDocument.Parse(texto); } catch (JsonException) { }
            if (respuesta.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
            { doc?.Dispose(); return ResultadoBugzilla<JsonDocument>.Fallo("Bugzilla rechazó la credencial o los permisos de la cuenta de servicio."); }
            if (doc is not null && doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("error", out var e) && e.ValueKind == JsonValueKind.True)
            {
                var mensaje = doc.RootElement.TryGetProperty("message", out var m) && m.ValueKind == JsonValueKind.String ? m.GetString() : null;
                doc.Dispose();
                return ResultadoBugzilla<JsonDocument>.Fallo("Bugzilla: " + Acotar(mensaje ?? $"error {(int)respuesta.StatusCode}"));
            }
            if (!respuesta.IsSuccessStatusCode || doc is null)
            { doc?.Dispose(); return ResultadoBugzilla<JsonDocument>.Fallo($"Bugzilla respondió {(int)respuesta.StatusCode}."); }
            return ResultadoBugzilla<JsonDocument>.Correcto(doc);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested) { return ResultadoBugzilla<JsonDocument>.Fallo("Bugzilla no respondió a tiempo."); }
        catch (HttpRequestException) { return ResultadoBugzilla<JsonDocument>.Fallo("No se pudo conectar con Bugzilla."); }
    }

    private static string Acotar(string texto)
    {
        var limpio = System.Text.RegularExpressions.Regex.Replace(texto, "<[^>]+>", " ");
        limpio = System.Text.RegularExpressions.Regex.Replace(limpio, @"\s+", " ").Trim();
        return limpio.Length <= 250 ? limpio : limpio[..250];
    }
}
