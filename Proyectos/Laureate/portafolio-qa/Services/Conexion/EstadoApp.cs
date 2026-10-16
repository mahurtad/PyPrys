namespace PortafolioQA.Services.Conexion;

public enum FaseApp
{
    /// <summary>Sin conexión configurada, o la base no tiene esquema: solo el asistente (con código) está disponible.</summary>
    Pendiente,
    /// <summary>Hay conexión y esquema, pero faltan migraciones (p. ej. tras actualizar): solo un Administrador usa el asistente.</summary>
    EsquemaDesactualizado,
    Activa
}

/// <summary>
/// Fase de la aplicación, evaluada al arrancar. Cualquier cambio de configuración o de esquema reinicia la aplicación,
/// por lo que no hace falta reevaluarla en caliente.
/// </summary>
public class EstadoApp
{
    public FaseApp Fase { get; private set; } = FaseApp.Pendiente;
    public string? Detalle { get; private set; }

    public async Task EvaluarAsync(ConexionStore store, ConexionProbador probador, ILogger logger)
    {
        if (store.Cadena is null)
        {
            Fase = FaseApp.Pendiente;
            Detalle = "No hay una conexión a la base de datos configurada.";
            return;
        }

        try
        {
            var (pendientes, aplicadas) = await probador.EstadoEsquemaAsync(store.Cadena);
            if (pendientes.Count == 0)
            {
                Fase = FaseApp.Activa;
            }
            else if (aplicadas == 0)
            {
                Fase = FaseApp.Pendiente;
                Detalle = "La base de datos no tiene el esquema de la aplicación.";
            }
            else
            {
                Fase = FaseApp.EsquemaDesactualizado;
                Detalle = $"Faltan {pendientes.Count} migraciones por aplicar a la base de datos.";
            }
        }
        catch (Exception ex)
        {
            // Una caída temporal de SQL no debe exponer el asistente: la aplicación sigue "activa" y /healthz informa el fallo.
            Fase = FaseApp.Activa;
            Detalle = "No se pudo comprobar el esquema al arrancar (sin conexión con la base de datos).";
            logger.LogError(ex, "No se pudo comprobar el esquema de la base de datos al arrancar.");
        }
    }
}
