namespace PortafolioQA.Services.Integrations;

public record SquashProyectoDto(int Id, string Nombre);
public record SquashCasoDto(int ItemId, string Nombre, string Estado, string? AsignadoLogin, DateTime? UltimaEjecucion, int? TestCaseId = null);
public record SquashModuloDto(int Id, string Nombre, List<int> ItemIds);
public record SquashIteracionDto(int Id, string Nombre, string? Estado, List<SquashCasoDto> Casos, List<SquashModuloDto>? Modulos = null);
public record SquashCampanaDto(int Id, string Nombre, string? Estado, List<SquashIteracionDto> Iteraciones);

public class ResultadoSquash<T>
{
    public bool Ok { get; init; }
    public string? Error { get; init; }
    public T? Datos { get; init; }

    public static ResultadoSquash<T> Correcto(T datos) => new() { Ok = true, Datos = datos };
    public static ResultadoSquash<T> Fallo(string error) => new() { Ok = false, Error = error };
}

/// <summary>Lecturas de Squash TM que necesita el portafolio. Solo lectura: nunca modifica Squash.</summary>
public interface ISquashApi
{
    Task<ResultadoSquash<List<SquashProyectoDto>>> ListarProyectosAsync(CancellationToken ct = default);

    /// <summary>Campañas del proyecto de Squash, con sus iteraciones y los casos de cada plan de pruebas.</summary>
    Task<ResultadoSquash<List<SquashCampanaDto>>> ObtenerEjecucionAsync(int squashProjectId, CancellationToken ct = default);
}

public class SquashNoConfigurado : ISquashApi
{
    private const string Mensaje = "Squash TM no está configurado (falta el token en Squash:Token).";

    public Task<ResultadoSquash<List<SquashProyectoDto>>> ListarProyectosAsync(CancellationToken ct = default) =>
        Task.FromResult(ResultadoSquash<List<SquashProyectoDto>>.Fallo(Mensaje));

    public Task<ResultadoSquash<List<SquashCampanaDto>>> ObtenerEjecucionAsync(int squashProjectId, CancellationToken ct = default) =>
        Task.FromResult(ResultadoSquash<List<SquashCampanaDto>>.Fallo(Mensaje));
}

/// <summary>
/// Porcentaje de avance de un caso según su estado en Squash. Null = no cuenta (no ejecutable, omitido, cancelado).
/// Sin ejecutar y bloqueado = 0; en curso, fallido, con error o advertencia = 50; aprobado o corregido = 100.
/// </summary>
public static class SquashEstados
{
    public static decimal? Porcentaje(string? estado) => estado switch
    {
        "SUCCESS" or "SETTLED" => 100m,
        "RUNNING" or "FAILURE" or "ERROR" or "WARNING" => 50m,
        "UNTESTABLE" or "SKIPPED" or "CANCELLED" or "NOT_FOUND" => null,
        _ => 0m // READY, BLOCKED, NOT_RUN y cualquier estado desconocido
    };

    public static string Texto(string? estado) => estado switch
    {
        "READY" => "Sin ejecutar",
        "RUNNING" => "En curso",
        "SUCCESS" => "Aprobado",
        "FAILURE" => "Fallido",
        "BLOCKED" => "Bloqueado",
        "ERROR" => "Con error",
        "WARNING" => "Advertencia",
        "SETTLED" => "Corregido",
        "UNTESTABLE" => "No ejecutable",
        "SKIPPED" => "Omitido",
        "CANCELLED" => "Cancelado",
        "NOT_RUN" => "No ejecutado",
        "NOT_FOUND" => "No encontrado",
        _ => estado ?? "—"
    };

    public static string Badge(string? estado) => estado switch
    {
        "SUCCESS" or "SETTLED" => "bg-success",
        "RUNNING" => "bg-primary",
        "FAILURE" or "ERROR" => "bg-danger",
        "WARNING" => "bg-warning text-dark",
        "BLOCKED" => "bg-dark",
        _ => "bg-secondary"
    };
}
