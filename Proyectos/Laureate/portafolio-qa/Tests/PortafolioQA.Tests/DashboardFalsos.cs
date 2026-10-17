using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Tests;

public class DefectosFalsos : IIntegracionDefectos, IVerificaProductoDefectos
{
    public Dictionary<string, ResultadoDefectos> PorProducto { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, int> Llamadas { get; } = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _lock = new();
    public int EnCurso, MaximoSimultaneo;
    public HashSet<string> Inexistentes { get; } = new(StringComparer.OrdinalIgnoreCase);

    public Task<bool?> ProductoExisteAsync(string producto, CancellationToken ct = default) => Task.FromResult<bool?>(!Inexistentes.Contains(producto));

    public async Task<ResultadoDefectos> ObtenerDefectosAsync(string producto, CancellationToken ct = default)
    {
        lock (_lock) { Llamadas[producto] = Llamadas.GetValueOrDefault(producto) + 1; EnCurso++; MaximoSimultaneo = Math.Max(MaximoSimultaneo, EnCurso); }
        await Task.Delay(50, ct);
        lock (_lock) EnCurso--;
        return PorProducto.TryGetValue(producto, out var r) ? r : new ResultadoDefectos { Disponible = false, Mensaje = "Bugzilla no pudo consultar el producto." };
    }

    public static DefectoExterno Bug(int id, string estado, string severidad, string? resolucion = null) =>
        new() { Id = id.ToString(), Titulo = "Bug " + id, Estado = estado, Severidad = severidad, Resolucion = resolucion, Url = "http://x/" + id };
}

public class RelojFijo(DateTime ahora) : TimeProvider
{
    public override DateTimeOffset GetUtcNow() => new(DateTime.SpecifyKind(ahora, DateTimeKind.Utc));
    public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
}
