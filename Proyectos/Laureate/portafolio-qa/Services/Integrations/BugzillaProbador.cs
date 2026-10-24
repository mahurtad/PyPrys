using System.Text.Json;

namespace PortafolioQA.Services.Integrations;

/// <param name="TienePermiso">null si Bugzilla no informó los grupos de la cuenta: no se puede afirmar ni negar editcomponents.</param>
/// <param name="RutaRest">"rest" o "rest.cgi": la que respondió en esa dirección (solo cuando la dirección es Bugzilla).</param>
public record ResultadoPruebaBugzilla(bool Conecta, bool ClaveValida, bool CuentaCoincide, bool? TienePermiso, string Mensaje, string? RutaRest = null)
{
    /// <summary>Se puede guardar: la clave es válida, es de esa cuenta y no consta que falte editcomponents.</summary>
    public bool Aceptable => ClaveValida && CuentaCoincide && TienePermiso != false;
}

/// <summary>
/// Valida una clave API de la cuenta de servicio sin crear nada en Bugzilla: (1) la clave autentica, (2) pertenece al login
/// indicado y (3) esa cuenta pertenece al grupo editcomponents. Bugzilla 5.2 no tiene /whoami; se usan valid_login y user.
/// La clave viaja como parámetro Bugzilla_api_key (la cabecera se ignora) y nunca aparece en mensajes ni en el registro.
/// </summary>
public class BugzillaProbador(HttpClient http)
{
    public const string GrupoRequerido = "editcomponents";

    public async Task<ResultadoPruebaBugzilla> ProbarAsync(string? baseUrl, string? apiKey, string? cuenta, CancellationToken ct = default)
    {
        if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https"))
            return Falla("Bugzilla:BaseUrl no está configurada en el servidor.", conecta: false);
        if (string.IsNullOrWhiteSpace(apiKey)) return Falla("Pegue la clave API.");
        if (string.IsNullOrWhiteSpace(cuenta)) return Falla("Indique el login de la cuenta de servicio.");

        var baseN = BugzillaCredencialesStore.NormalizarUrl(baseUrl);
        var clave = Uri.EscapeDataString(apiKey.Trim());
        var login = Uri.EscapeDataString(cuenta.Trim());
        try
        {
            var ruta = await DetectarRutaAsync(baseN, ct);
            if (ruta is null) return NoEsBugzilla;
            var raiz = $"{baseN}/{ruta}/";
            using var v = await ObtenerAsync($"{raiz}valid_login?login={login}&Bugzilla_api_key={clave}", ct);
            if (v is null) return Falla("Bugzilla respondió algo que no se pudo interpretar. ¿La dirección es la de Bugzilla?");
            if (v.RootElement.TryGetProperty("error", out var err) && err.ValueKind == JsonValueKind.True)
                return Falla(Codigo(v.RootElement) == 306
                    ? "Bugzilla no reconoce esa clave API. Revise que la copió completa."
                    : "Bugzilla rechazó la consulta: " + Mensaje(v.RootElement));
            if (!(v.RootElement.TryGetProperty("result", out var res) && res.ValueKind == JsonValueKind.True))
                return new(true, true, false, null, "La clave es válida pero no pertenece a la cuenta " + cuenta.Trim() +
                    ". Genere la clave iniciando sesión en Bugzilla con esa cuenta.");

            using var u = await ObtenerAsync($"{raiz}user?names={login}&Bugzilla_api_key={clave}", ct);
            bool? permiso = null;
            if (u is not null && u.RootElement.TryGetProperty("users", out var users) && users.ValueKind == JsonValueKind.Array
                && users.GetArrayLength() > 0 && users[0].TryGetProperty("groups", out var grupos) && grupos.ValueKind == JsonValueKind.Array)
                permiso = grupos.EnumerateArray().Any(g => g.TryGetProperty("name", out var n)
                    && string.Equals(n.GetString(), GrupoRequerido, StringComparison.OrdinalIgnoreCase));

            ResultadoPruebaBugzilla resultado = permiso switch
            {
                true => new(true, true, true, true, "Conexión correcta: la clave es de esa cuenta y tiene el grupo editcomponents."),
                false => new(true, true, true, false, "La clave es de esa cuenta, pero la cuenta no tiene el grupo editcomponents. " +
                    "Pida al administrador de Bugzilla que se lo asigne (y solo ese grupo)."),
                _ => new(true, true, true, null, "La clave es de esa cuenta, pero Bugzilla no informó sus grupos: no se pudo comprobar " +
                    "editcomponents. Si falta, los productos quedarán en Error al crearse.")
            };
            return resultado with { RutaRest = ruta };
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested) { return Falla("Bugzilla no respondió a tiempo.", conecta: false); }
        catch (HttpRequestException) { return Falla("No se pudo conectar con Bugzilla.", conecta: false); }
    }

    /// <summary>
    /// Prueba la clave de lectura (consulta de defectos). Solo comprueba que Bugzilla la reconoce: sin conocer el login de esa
    /// cuenta no se puede afirmar a quién pertenece.
    /// </summary>
    public async Task<ResultadoPruebaBugzilla> ProbarLecturaAsync(string? baseUrl, string? apiKey, CancellationToken ct = default)
    {
        if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https"))
            return Falla("Indique la dirección de Bugzilla.", conecta: false);
        if (string.IsNullOrWhiteSpace(apiKey)) return Falla("Pegue la clave de lectura.");
        var baseN = BugzillaCredencialesStore.NormalizarUrl(baseUrl);
        try
        {
            var ruta = await DetectarRutaAsync(baseN, ct);
            if (ruta is null) return NoEsBugzilla;
            var raiz = $"{baseN}/{ruta}/";
            using var p = await ObtenerAsync($"{raiz}product?type=accessible&include_fields=id&Bugzilla_api_key={Uri.EscapeDataString(apiKey.Trim())}", ct);
            if (p is null) return Falla("Bugzilla respondió algo que no se pudo interpretar.");
            if (p.RootElement.TryGetProperty("error", out var err) && err.ValueKind == JsonValueKind.True)
                return Falla(Codigo(p.RootElement) == 306 ? "Bugzilla no reconoce la clave de lectura. Revise que la copió completa."
                    : "Bugzilla rechazó la consulta: " + Mensaje(p.RootElement));
            return new(true, true, true, null, "La clave de lectura es reconocida por Bugzilla.", ruta);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested) { return Falla("Bugzilla no respondió a tiempo.", conecta: false); }
        catch (HttpRequestException) { return Falla("No se pudo conectar con Bugzilla.", conecta: false); }
    }

    /// <summary>Rutas de la API REST que se prueban, en orden: con reescritura de Apache y la que siempre existe (rest.cgi).</summary>
    public static readonly string[] RutasRest = { "rest", "rest.cgi" };

    private static readonly ResultadoPruebaBugzilla NoEsBugzilla = Falla(
        "La dirección no responde como Bugzilla (ni en /rest ni en /rest.cgi). Revise la dirección: es donde está instalado Bugzilla, " +
        "sin /rest; por ejemplo https://bugzilla.su-dominio o http://servidor/bugzilla.", conecta: false);

    /// <summary>La ruta REST que responde <c>version</c> en esa dirección, o null si ninguna lo hace. No se envía ninguna clave.</summary>
    private async Task<string?> DetectarRutaAsync(string baseN, CancellationToken ct)
    {
        foreach (var ruta in RutasRest)
        {
            using var v = await ObtenerAsync($"{baseN}/{ruta}/version", ct);
            if (v is not null && v.RootElement.ValueKind == JsonValueKind.Object && v.RootElement.TryGetProperty("version", out _)) return ruta;
        }
        return null;
    }

    private async Task<JsonDocument?> ObtenerAsync(string url, CancellationToken ct)
    {
        using var respuesta = await http.GetAsync(url, ct);
        var texto = await respuesta.Content.ReadAsStringAsync(ct);
        try { return JsonDocument.Parse(texto); }
        catch (JsonException) { return null; }
    }

    private static int Codigo(JsonElement raiz) => raiz.TryGetProperty("code", out var c) && c.TryGetInt32(out var n) ? n : 0;

    private static string Mensaje(JsonElement raiz)
    {
        var m = raiz.TryGetProperty("message", out var e) && e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : "";
        m = System.Text.RegularExpressions.Regex.Replace(m, "<[^>]+>", " ");
        m = System.Text.RegularExpressions.Regex.Replace(m, @"\s+", " ").Trim();
        return m.Length <= 200 ? m : m[..200];
    }

    private static ResultadoPruebaBugzilla Falla(string mensaje, bool conecta = true) => new(conecta, false, false, null, mensaje);
}
