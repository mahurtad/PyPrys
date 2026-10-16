using System.Net;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace PortafolioQA.Services.Integrations;

public class SquashOptions
{
    /// <summary>Raíz de Squash TM, p. ej. http://servidor:8090/squash (sin /api/rest).</summary>
    public string? BaseUrl { get; set; }

    /// <summary>Token de API (user-secrets). Para la integración basta con permiso de lectura.</summary>
    public string? Token { get; set; }

    public bool Configurado => !string.IsNullOrWhiteSpace(BaseUrl) && !string.IsNullOrWhiteSpace(Token);
}

/// <summary>
/// Cliente de solo lectura de la API REST de Squash TM. El identificador externo es "campaign:ID" o "iteration:ID".
/// Planificados = ítems del plan de pruebas; Pendientes = READY o RUNNING; Ejecutados = Planificados − Pendientes.
/// Nunca devuelve ceros inventados: si Squash no responde o no hay ítems, Disponible = false con el motivo.
/// </summary>
public class SquashTmClient : IIntegracionCertificacion, ISquashApi
{
    private const int TamanoPagina = 200;
    private readonly HttpClient _http;

    public SquashTmClient(HttpClient http)
    {
        _http = http;
    }

    public static bool TryParseIdentificador(string? texto, out string tipo, out int id)
    {
        tipo = string.Empty; id = 0;
        var partes = (texto ?? string.Empty).Trim().Split(':');
        if (partes.Length != 2 || !int.TryParse(partes[1], out id) || id <= 0) return false;
        tipo = partes[0].Trim().ToLowerInvariant();
        return tipo is "campaign" or "iteration";
    }

    public async Task<ResultadoCertificacion> ObtenerAvanceAsync(string identificadorExterno, CancellationToken cancellationToken = default)
    {
        if (!TryParseIdentificador(identificadorExterno, out var tipo, out var id))
            return NoDisponible("El identificador debe tener el formato campaign:ID o iteration:ID.");

        try
        {
            var iteraciones = new List<int>();
            if (tipo == "iteration") iteraciones.Add(id);
            else
            {
                using var campana = await GetAsync($"campaigns/{id}", cancellationToken);
                if (campana.RootElement.TryGetProperty("iterations", out var its) && its.ValueKind == JsonValueKind.Array)
                    iteraciones.AddRange(its.EnumerateArray().Select(i => i.GetProperty("id").GetInt32()));
            }

            int planificados = 0, pendientes = 0, aprobados = 0, fallidos = 0, bloqueados = 0;
            foreach (var iteracion in iteraciones)
            {
                for (var pagina = 0; ; pagina++)
                {
                    using var doc = await GetAsync($"iterations/{iteracion}/test-plan?size={TamanoPagina}&page={pagina}", cancellationToken);
                    var raiz = doc.RootElement;
                    if (raiz.TryGetProperty("_embedded", out var emb) && emb.TryGetProperty("test-plan", out var items))
                    {
                        foreach (var item in items.EnumerateArray())
                        {
                            planificados++;
                            switch (item.TryGetProperty("execution_status", out var s) ? s.GetString() : null)
                            {
                                case "READY": case "RUNNING": case null: pendientes++; break;
                                case "SUCCESS": case "SETTLED": aprobados++; break;
                                case "FAILURE": case "ERROR": fallidos++; break;
                                case "BLOCKED": bloqueados++; break;
                                // UNTESTABLE, SKIPPED, WARNING…: cuentan como ejecutados, sin categoría propia.
                            }
                        }
                    }
                    var totalPaginas = raiz.TryGetProperty("page", out var pg) && pg.TryGetProperty("totalPages", out var tp) ? tp.GetInt32() : 1;
                    if (pagina + 1 >= totalPaginas) break;
                }
            }

            if (planificados == 0) return NoDisponible("La campaña de Squash TM no tiene casos en su plan de pruebas.");
            return new ResultadoCertificacion
            {
                Disponible = true,
                ElementosPlanificados = planificados,
                ElementosEjecutados = planificados - pendientes,
                Aprobados = aprobados,
                Fallidos = fallidos,
                Bloqueados = bloqueados,
                Pendientes = pendientes,
                UltimaActualizacion = DateTime.Now
            };
        }
        catch (SquashException ex) { return NoDisponible(ex.Message); }
        catch (HttpRequestException) { return NoDisponible("No se pudo conectar con Squash TM."); }
        catch (TaskCanceledException) when (!cancellationToken.IsCancellationRequested) { return NoDisponible("Squash TM no respondió a tiempo."); }
        catch (JsonException) { return NoDisponible("Squash TM devolvió una respuesta que no se pudo interpretar."); }
    }

    public async Task<ResultadoSquash<List<SquashProyectoDto>>> ListarProyectosAsync(CancellationToken ct = default)
    {
        try
        {
            var lista = new List<SquashProyectoDto>();
            await foreach (var p in PaginarAsync("projects", "projects", ct))
                lista.Add(new SquashProyectoDto(p.GetProperty("id").GetInt32(), p.GetProperty("name").GetString() ?? string.Empty));
            return ResultadoSquash<List<SquashProyectoDto>>.Correcto(lista.OrderBy(p => p.Nombre).ToList());
        }
        catch (Exception ex) when (EsErrorDeSquash(ex)) { return ResultadoSquash<List<SquashProyectoDto>>.Fallo(Describir(ex)); }
    }

    public async Task<ResultadoSquash<List<SquashCampanaDto>>> ObtenerEjecucionAsync(int squashProjectId, CancellationToken ct = default)
    {
        try
        {
            var campanas = new List<SquashCampanaDto>();
            await foreach (var c in PaginarAsync($"projects/{squashProjectId}/campaigns", "campaigns", ct))
            {
                var campanaId = c.GetProperty("id").GetInt32();
                using var detalle = await GetAsync($"campaigns/{campanaId}", ct);
                var raiz = detalle.RootElement;
                var iteraciones = new List<SquashIteracionDto>();
                if (raiz.TryGetProperty("iterations", out var its) && its.ValueKind == JsonValueKind.Array)
                {
                    foreach (var it in its.EnumerateArray())
                    {
                        var iteracionId = it.GetProperty("id").GetInt32();
                        var casos = new List<SquashCasoDto>();
                        await foreach (var item in PaginarAsync($"iterations/{iteracionId}/test-plan", "test-plan", ct))
                            casos.Add(LeerCaso(item));
                        iteraciones.Add(new SquashIteracionDto(iteracionId, Texto(it, "name") ?? $"Iteración {iteracionId}", null, casos));
                    }
                }
                campanas.Add(new SquashCampanaDto(campanaId, Texto(raiz, "name") ?? $"Campaña {campanaId}", Texto(raiz, "progress_status"), iteraciones));
            }
            return ResultadoSquash<List<SquashCampanaDto>>.Correcto(campanas);
        }
        catch (Exception ex) when (EsErrorDeSquash(ex)) { return ResultadoSquash<List<SquashCampanaDto>>.Fallo(Describir(ex)); }
    }

    private static SquashCasoDto LeerCaso(JsonElement item)
    {
        var id = item.GetProperty("id").GetInt32();
        var nombre = item.TryGetProperty("referenced_test_case", out var tc) && tc.ValueKind == JsonValueKind.Object ? Texto(tc, "name") : null;
        int? tcId = item.TryGetProperty("referenced_test_case", out var tcRef) && tcRef.ValueKind == JsonValueKind.Object
            && tcRef.TryGetProperty("id", out var tcIdEl) && tcIdEl.TryGetInt32(out var tcIdVal) ? tcIdVal : null;
        string? login = null;
        if (item.TryGetProperty("assigned_to", out var asignado))
            login = asignado.ValueKind == JsonValueKind.Object ? Texto(asignado, "login") : asignado.ValueKind == JsonValueKind.String ? asignado.GetString() : null;
        DateTime? ultima = item.TryGetProperty("last_executed_on", out var fecha) && fecha.ValueKind == JsonValueKind.String
            && DateTimeOffset.TryParse(fecha.GetString(), out var f) ? f.LocalDateTime : null;
        return new SquashCasoDto(id, nombre ?? $"Ítem {id}", Texto(item, "execution_status") ?? "READY", login, ultima, tcId);
    }

    private static string? Texto(JsonElement elemento, string propiedad) =>
        elemento.TryGetProperty(propiedad, out var p) && p.ValueKind == JsonValueKind.String ? p.GetString() : null;

    private async IAsyncEnumerable<JsonElement> PaginarAsync(string ruta, string clave, [EnumeratorCancellation] CancellationToken ct)
    {
        var separador = ruta.Contains('?') ? '&' : '?';
        for (var pagina = 0; ; pagina++)
        {
            using var doc = await GetAsync($"{ruta}{separador}size={TamanoPagina}&page={pagina}", ct);
            var raiz = doc.RootElement;
            if (raiz.TryGetProperty("_embedded", out var emb) && emb.TryGetProperty(clave, out var items))
                foreach (var item in items.EnumerateArray()) yield return item.Clone();
            var totalPaginas = raiz.TryGetProperty("page", out var pg) && pg.TryGetProperty("totalPages", out var tp) ? tp.GetInt32() : 1;
            if (pagina + 1 >= totalPaginas) yield break;
        }
    }

    private static bool EsErrorDeSquash(Exception ex) =>
        ex is SquashException or HttpRequestException or TaskCanceledException or JsonException or KeyNotFoundException or InvalidOperationException;

    private static string Describir(Exception ex) => ex switch
    {
        SquashException s => s.Message,
        HttpRequestException => "No se pudo conectar con Squash TM.",
        TaskCanceledException => "Squash TM no respondió a tiempo.",
        _ => "Squash TM devolvió una respuesta que no se pudo interpretar."
    };

    private async Task<JsonDocument> GetAsync(string ruta, CancellationToken ct)
    {
        using var respuesta = await _http.GetAsync(ruta, ct);
        if (respuesta.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
            throw new SquashException("Squash TM rechazó el token (401/403): revisa el token configurado.");
        if (respuesta.StatusCode == HttpStatusCode.NotFound)
            throw new SquashException("No existe en Squash TM el elemento indicado en el identificador externo.");
        if (!respuesta.IsSuccessStatusCode)
            throw new SquashException($"Squash TM respondió {(int)respuesta.StatusCode}.");
        return await JsonDocument.ParseAsync(await respuesta.Content.ReadAsStreamAsync(ct), cancellationToken: ct);
    }

    private static ResultadoCertificacion NoDisponible(string mensaje) => new() { Disponible = false, Mensaje = mensaje };

    private sealed class SquashException : Exception
    {
        public SquashException(string mensaje) : base(mensaje) { }
    }
}
