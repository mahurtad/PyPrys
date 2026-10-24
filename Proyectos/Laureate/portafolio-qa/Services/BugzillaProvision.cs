using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Services;

public record ResultadoPreparacion(bool Aplica, EstadoBugzilla? Estado, string Mensaje);

/// <summary>
/// Prepara en Bugzilla (solo por su API REST) el producto de una iniciativa vinculada a Squash TM: producto con el nombre del
/// proyecto de Squash, componente "General" y versión "1.0". Es idempotente y reanudable: cada paso se verifica antes de crearse,
/// el avance se guarda al terminar cada uno y un reintento continúa desde lo pendiente. No borra ni reasigna nada.
/// </summary>
public class BugzillaProvision(PortafolioQAContext db, IBugzillaProductos api, IOptions<SquashOptions> squash,
    IBugzillaServicioConfig bugzillaConfig, ILogger<BugzillaProvision> log)
{
    public const string ComponenteInicial = "General";
    public const string VersionInicial = "1.0";
    private static readonly TimeSpan Reserva = TimeSpan.FromMinutes(2);

    public static string NormalizarAmbiente(string? url) => (url ?? string.Empty).Trim().TrimEnd('/').ToLowerInvariant();

    /// <summary>Marca en la descripción del producto: es la única evidencia automática de que lo creó esta iniciativa.</summary>
    public static string Marca(string squashAmbiente, int squashProjectId) => $"[portafolio-qa squash={squashAmbiente}#{squashProjectId}]";

    public async Task<ResultadoPreparacion> PrepararAsync(int proyectoId, bool confirmarExistente = false, CancellationToken ct = default)
    {
        var proyecto = await db.Proyectos.AsNoTracking().Where(p => p.Id == proyectoId)
            .Select(p => new { p.Id, p.Nombre, p.SquashProjectId }).FirstOrDefaultAsync(ct);
        if (proyecto is null) return new(false, null, "El proyecto no existe.");
        if (proyecto.SquashProjectId is not int squashId)
            return new(false, null, "El proyecto no está vinculado a Squash TM: no se prepara producto en Bugzilla.");

        var squashAmbiente = NormalizarAmbiente(squash.Value.BaseUrl);
        var bugzillaAmbiente = NormalizarAmbiente(bugzillaConfig.Actual.BaseUrl);

        var vinculo = await db.VinculosBugzilla.FirstOrDefaultAsync(v => v.ProyectoId == proyectoId, ct);
        if (vinculo is null)
        {
            vinculo = new VinculoBugzilla { ProyectoId = proyectoId, SquashAmbiente = squashAmbiente, SquashProjectId = squashId, BugzillaAmbiente = bugzillaAmbiente };
            db.VinculosBugzilla.Add(vinculo);
            try { await db.SaveChangesAsync(ct); }
            catch (DbUpdateException)
            {
                // Solicitud concurrente: otra ya creó el vínculo de esta iniciativa (índice único por ProyectoId).
                db.Entry(vinculo).State = EntityState.Detached;
                vinculo = await db.VinculosBugzilla.FirstOrDefaultAsync(v => v.ProyectoId == proyectoId, ct);
                if (vinculo is null)
                    return new(true, EstadoBugzilla.Error, "Ese proyecto de Squash TM ya tiene un vínculo de Bugzilla con otra iniciativa; no se reasigna automáticamente.");
            }
        }

        if (!bugzillaConfig.Actual.ProductosConfigurado)
            return await TerminarAsync(vinculo, EstadoBugzilla.Pendiente, BugzillaProductosNoConfigurado.Mensaje, ct);

        // Reserva atómica en la base: solo una solicitud a la vez prepara esta iniciativa (reintentos o clics repetidos).
        var ahora = DateTime.UtcNow;
        var reservado = await db.VinculosBugzilla.Where(v => v.Id == vinculo.Id && (v.BloqueadoHasta == null || v.BloqueadoHasta < ahora))
            .ExecuteUpdateAsync(s => s.SetProperty(v => v.BloqueadoHasta, ahora + Reserva), ct);
        if (reservado == 0)
            return new(true, vinculo.Estado, "Ya hay una preparación en curso para esta iniciativa; espera unos segundos y reintenta.");

        try
        {
            await db.Entry(vinculo).ReloadAsync(ct);

            // Con producto ya creado, el vínculo solo vale en el ambiente donde se creó. Sin producto aún, se alinea al ambiente actual.
            var tieneProducto = vinculo.ProductoId is not null || vinculo.ComponenteId is not null;
            if (vinculo.SquashAmbiente != squashAmbiente || vinculo.SquashProjectId != squashId || (tieneProducto && vinculo.BugzillaAmbiente != bugzillaAmbiente))
            {
                if (tieneProducto)
                    return await TerminarAsync(vinculo, EstadoBugzilla.Error,
                        "El vínculo guardado pertenece a otro ambiente o a otro proyecto de Squash TM. No se reutiliza ni se reasigna automáticamente; revísalo con un administrador.", ct);
                vinculo.SquashAmbiente = squashAmbiente;
                vinculo.SquashProjectId = squashId;
                vinculo.BugzillaAmbiente = bugzillaAmbiente;
            }
            if (confirmarExistente) vinculo.AsociacionExistenteConfirmada = true;
            return await EjecutarPasosAsync(vinculo, proyecto.Nombre, ct);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            log.LogError(ex, "Falló la preparación del producto Bugzilla del proyecto {ProyectoId}.", proyectoId);
            db.ChangeTracker.Clear();
            var fila = await db.VinculosBugzilla.FirstAsync(v => v.ProyectoId == proyectoId, CancellationToken.None);
            return await TerminarAsync(fila, EstadoBugzilla.Error, "Error inesperado al preparar Bugzilla; reintenta o revisa el registro de la aplicación.", CancellationToken.None);
        }
        finally
        {
            await db.VinculosBugzilla.Where(v => v.Id == vinculo.Id)
                .ExecuteUpdateAsync(s => s.SetProperty(v => v.BloqueadoHasta, (DateTime?)null), CancellationToken.None);
        }
    }

    private async Task<ResultadoPreparacion> EjecutarPasosAsync(VinculoBugzilla v, string nombreProyecto, CancellationToken ct)
    {
        var marca = Marca(v.SquashAmbiente, v.SquashProjectId);
        var descripcion = $"Producto de la iniciativa «{nombreProyecto}» creado por Portafolio QA. {marca}";

        // 1. Producto: reutiliza el vínculo guardado (por ID, no por nombre); si no hay, busca o crea.
        ProductoBugzilla? producto;
        if (v.ProductoId is int productoId)
        {
            var lectura = await api.ObtenerProductoAsync(productoId, ct);
            if (!lectura.Ok) return await TerminarAsync(v, EstadoBugzilla.Error, "No se pudo verificar el producto vinculado. " + lectura.Error, ct);
            producto = lectura.Datos;
            if (producto is null)
                return await TerminarAsync(v, EstadoBugzilla.Error, "El producto vinculado ya no existe en Bugzilla. No se crea otro automáticamente; revísalo con un administrador.", ct);
        }
        else
        {
            var nombre = nombreProyecto.Trim();
            var lista = await api.ListarProductosAsync(ct);
            if (!lista.Ok) return await TerminarAsync(v, EstadoBugzilla.Error, "No se pudo consultar Bugzilla. " + lista.Error, ct);
            var existente = Buscar(lista.Datos!, nombre);
            if (existente is null)
            {
                var creado = await api.CrearProductoAsync(nombre, descripcion, VersionInicial, ct);
                if (!creado.Ok)
                {
                    // Puede que otra solicitud lo haya creado justo antes: se vuelve a consultar antes de declarar el error.
                    var relista = await api.ListarProductosAsync(ct);
                    existente = relista.Ok ? Buscar(relista.Datos!, nombre) : null;
                    if (existente is null) return await TerminarAsync(v, EstadoBugzilla.Error, "No se pudo crear el producto. " + creado.Error, ct);
                }
                else existente = new ProductoResumenBugzilla(creado.Datos, nombre);
            }
            var detalle = await api.ObtenerProductoAsync(existente.Id, ct);
            if (!detalle.Ok || detalle.Datos is null)
                return await TerminarAsync(v, EstadoBugzilla.Error, "No se pudo verificar el producto. " + (detalle.Error ?? "No existe."), ct);
            producto = detalle.Datos;

            // Un nombre coincidente no basta: debe ser un producto que esta iniciativa creó (marca) o que un administrador confirmó.
            if (!(producto.Descripcion ?? string.Empty).Contains(marca, StringComparison.Ordinal) && !v.AsociacionExistenteConfirmada)
                return await TerminarAsync(v, EstadoBugzilla.Error,
                    $"Ya existe en Bugzilla un producto llamado «{producto.Nombre}» que no se puede atribuir a esta iniciativa. Si corresponde, un administrador puede usar «Asociar existente».", ct);
            var ambiente = v.BugzillaAmbiente;
            if (await db.VinculosBugzilla.AnyAsync(x => x.Id != v.Id && x.BugzillaAmbiente == ambiente && x.ProductoId == producto.Id, ct))
                return await TerminarAsync(v, EstadoBugzilla.Error, "Ese producto de Bugzilla ya está vinculado a otra iniciativa.", ct);

            v.ProductoId = producto.Id;
            v.ProductoNombre = Recortar(producto.Nombre, 200);
            if (!await GuardarAsync(v, ct))
                return await TerminarAsync(v, EstadoBugzilla.Error, "Ese producto de Bugzilla ya está vinculado a otra iniciativa.", ct);
        }

        // 2. Componente "General": se crea solo si no existe (la comparación ignora mayúsculas, como Bugzilla).
        var componente = producto.Componentes.FirstOrDefault(c => string.Equals(c.Nombre, ComponenteInicial, StringComparison.OrdinalIgnoreCase));
        if (componente is null)
        {
            var creado = await api.CrearComponenteAsync(producto.Nombre, ComponenteInicial, "Componente general", bugzillaConfig.Actual.CuentaServicio, ct);
            if (!creado.Ok)
            {
                var releido = await api.ObtenerProductoAsync(producto.Id, ct);
                componente = releido.Datos?.Componentes.FirstOrDefault(c => string.Equals(c.Nombre, ComponenteInicial, StringComparison.OrdinalIgnoreCase));
                if (componente is null) return await TerminarAsync(v, EstadoBugzilla.Error, "El producto existe pero no se pudo crear el componente «General». " + creado.Error, ct);
            }
            else componente = new ComponenteBugzilla(creado.Datos, ComponenteInicial);
            v.ComponenteId = componente.Id; v.ComponenteNombre = ComponenteInicial;
            await GuardarAsync(v, ct);
        }

        // 3. Verificación final de producto, componente y versión tal como quedaron en Bugzilla.
        var final = await api.ObtenerProductoAsync(producto.Id, ct);
        if (!final.Ok || final.Datos is null)
            return await TerminarAsync(v, EstadoBugzilla.Error, "No se pudo verificar el resultado en Bugzilla. " + (final.Error ?? "El producto no existe."), ct);
        var comp = final.Datos.Componentes.FirstOrDefault(c => string.Equals(c.Nombre, ComponenteInicial, StringComparison.OrdinalIgnoreCase));
        if (comp is null) return await TerminarAsync(v, EstadoBugzilla.Error, "Falta el componente «General» en el producto de Bugzilla.", ct);
        v.ComponenteId = comp.Id; v.ComponenteNombre = comp.Nombre;
        var version = final.Datos.Versiones.FirstOrDefault(x => string.Equals(x.Nombre, VersionInicial, StringComparison.OrdinalIgnoreCase));
        if (version is null)
        {
            // Bugzilla 5.2 solo crea la versión al crear el producto; su API REST no permite agregar versiones a uno existente.
            await GuardarAsync(v, ct);
            return await TerminarAsync(v, EstadoBugzilla.Error, "Falta la versión «1.0» y la API de Bugzilla 5.2 no permite crearla: agrégala manualmente en el producto y reintenta.", ct);
        }
        v.VersionId = version.Id; v.VersionNombre = version.Nombre;
        v.UltimaSincronizacion = DateTime.Now;
        return await TerminarAsync(v, EstadoBugzilla.Configurado, "Producto, componente y versión verificados en Bugzilla.", ct);
    }

    private static ProductoResumenBugzilla? Buscar(List<ProductoResumenBugzilla> lista, string nombre) =>
        lista.FirstOrDefault(p => string.Equals(p.Nombre.Trim(), nombre, StringComparison.OrdinalIgnoreCase));

    private async Task<bool> GuardarAsync(VinculoBugzilla v, CancellationToken ct)
    {
        try { await db.SaveChangesAsync(ct); return true; }
        catch (DbUpdateException) { return false; }
    }

    private async Task<ResultadoPreparacion> TerminarAsync(VinculoBugzilla v, EstadoBugzilla estado, string mensaje, CancellationToken ct)
    {
        v.Estado = estado;
        v.UltimoIntento = DateTime.Now;
        v.UltimoError = estado == EstadoBugzilla.Configurado ? null : Recortar(mensaje, 500);
        try { await db.SaveChangesAsync(ct); }
        catch (DbUpdateException ex) { log.LogError(ex, "No se pudo guardar el estado de Bugzilla del vínculo {Id}.", v.Id); }
        return new(true, estado, mensaje);
    }

    private static string Recortar(string texto, int max) => texto.Length <= max ? texto : texto[..max];
}
