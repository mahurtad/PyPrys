using System.Net;
using System.Text;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using PortafolioQA.Data;
using PortafolioQA.Models;
using PortafolioQA.Services;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Tests;

/// <summary>Bugzilla simulado en memoria: mismas reglas que la API (nombres únicos sin distinguir mayúsculas) y fallos inyectables.</summary>
public class BugzillaFalso : IBugzillaProductos
{
    private int _siguiente = 100;
    private readonly object _lock = new();
    public List<(int Id, string Nombre, string Descripcion, List<string> Componentes, List<string> Versiones)> Productos { get; } = new();
    public int ProductosCreados, ComponentesCreados;
    public bool FallarProducto, FallarComponente, SinVersionPorApi;
    public TimeSpan Demora = TimeSpan.Zero;

    public async Task<ResultadoBugzilla<List<ProductoResumenBugzilla>>> ListarProductosAsync(CancellationToken ct = default)
    {
        await Task.Delay(Demora, ct);
        lock (_lock) return ResultadoBugzilla<List<ProductoResumenBugzilla>>.Correcto(Productos.Select(p => new ProductoResumenBugzilla(p.Id, p.Nombre)).ToList());
    }

    public async Task<ResultadoBugzilla<ProductoBugzilla?>> ObtenerProductoAsync(int id, CancellationToken ct = default)
    {
        await Task.Delay(Demora, ct);
        lock (_lock)
        {
            var p = Productos.FirstOrDefault(x => x.Id == id);
            if (p.Nombre is null) return ResultadoBugzilla<ProductoBugzilla?>.Correcto(null);
            return ResultadoBugzilla<ProductoBugzilla?>.Correcto(new ProductoBugzilla(p.Id, p.Nombre, p.Descripcion,
                p.Componentes.Select((c, i) => new ComponenteBugzilla(id * 10 + i, c)).ToList(),
                p.Versiones.Select((v, i) => new VersionBugzilla(id * 10 + i, v)).ToList()));
        }
    }

    public async Task<ResultadoBugzilla<int>> CrearProductoAsync(string nombre, string descripcion, string versionInicial, CancellationToken ct = default)
    {
        await Task.Delay(Demora, ct);
        lock (_lock)
        {
            if (FallarProducto) return ResultadoBugzilla<int>.Fallo("Bugzilla: error simulado al crear el producto.");
            if (Productos.Any(p => string.Equals(p.Nombre, nombre, StringComparison.OrdinalIgnoreCase)))
                return ResultadoBugzilla<int>.Fallo("Bugzilla: el producto ya existe.");
            var id = _siguiente++;
            Productos.Add((id, nombre, descripcion, new List<string>(), SinVersionPorApi ? new List<string>() : new List<string> { versionInicial }));
            ProductosCreados++;
            return ResultadoBugzilla<int>.Correcto(id);
        }
    }

    public async Task<ResultadoBugzilla<int>> CrearComponenteAsync(string producto, string nombre, string descripcion, string asignadoPorDefecto, CancellationToken ct = default)
    {
        await Task.Delay(Demora, ct);
        lock (_lock)
        {
            if (FallarComponente) return ResultadoBugzilla<int>.Fallo("Bugzilla: error simulado al crear el componente.");
            var p = Productos.First(x => x.Nombre == producto);
            if (p.Componentes.Contains(nombre, StringComparer.OrdinalIgnoreCase)) return ResultadoBugzilla<int>.Fallo("Bugzilla: el componente ya existe.");
            p.Componentes.Add(nombre);
            ComponentesCreados++;
            return ResultadoBugzilla<int>.Correcto(p.Id * 10 + p.Componentes.Count - 1);
        }
    }
}

public class Entorno : IAsyncDisposable
{
    public string Squash = "http://localhost:8090/squash";
    public string Bugzilla = "http://localhost:8080";
    public string? AdminKey = "clave-de-prueba";
    public BugzillaFalso Api = new();
    private readonly string _cadena = $"Server=.\\SQLEXPRESS;Database=PortafolioQA_test_{Guid.NewGuid():N};Trusted_Connection=True;TrustServerCertificate=True;MultipleActiveResultSets=true";

    public PortafolioQAContext NuevoContexto() => new(new DbContextOptionsBuilder<PortafolioQAContext>().UseSqlServer(_cadena).Options);

    public BugzillaProvision Servicio(PortafolioQAContext db) => new(db, Api,
        Options.Create(new SquashOptions { BaseUrl = Squash, Token = "t" }),
        new BugzillaServicioFija(new BugzillaOptions { BaseUrl = Bugzilla, ApiKey = "k", AdminApiKey = AdminKey }),
        NullLogger<BugzillaProvision>.Instance);

    public static async Task<Entorno> CrearAsync()
    {
        var e = new Entorno();
        using var db = e.NuevoContexto();
        await db.Database.EnsureCreatedAsync();
        return e;
    }

    public async Task<int> NuevoProyectoAsync(string nombre, int? squashId)
    {
        using var db = NuevoContexto();
        var inst = await db.Instituciones.FirstOrDefaultAsync() ?? db.Instituciones.Add(new Institucion { Nombre = "Inst" }).Entity;
        var p = new Proyecto { Nombre = nombre, Institucion = inst, SquashProjectId = squashId };
        db.Proyectos.Add(p);
        await db.SaveChangesAsync();
        return p.Id;
    }

    public async Task<ResultadoPreparacion> PrepararAsync(int proyectoId, bool confirmar = false)
    {
        using var db = NuevoContexto();
        return await Servicio(db).PrepararAsync(proyectoId, confirmar);
    }

    public async Task<List<VinculoBugzilla>> VinculosAsync()
    {
        using var db = NuevoContexto();
        return await db.VinculosBugzilla.AsNoTracking().ToListAsync();
    }

    public async ValueTask DisposeAsync()
    {
        using var db = NuevoContexto();
        await db.Database.EnsureDeletedAsync();
    }
}

public class BugzillaProvisionTests
{
    [Fact]
    public async Task Crea_producto_componente_y_version_y_guarda_vinculo_con_ambientes()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);

        var r = await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Configurado, r.Estado);
        var p = Assert.Single(e.Api.Productos);
        Assert.Equal("Proyecto Prueba 3", p.Nombre);
        Assert.Equal(new[] { "General" }, p.Componentes);
        Assert.Equal(new[] { "1.0" }, p.Versiones);
        var v = Assert.Single(await e.VinculosAsync());
        Assert.Equal(p.Id, v.ProductoId);
        Assert.NotNull(v.ComponenteId); Assert.NotNull(v.VersionId);
        Assert.Equal("http://localhost:8090/squash", v.SquashAmbiente);
        Assert.Equal("http://localhost:8080", v.BugzillaAmbiente);
        Assert.Equal(3, v.SquashProjectId);
        Assert.NotNull(v.UltimaSincronizacion); Assert.Null(v.UltimoError);
    }

    [Fact]
    public async Task Reutiliza_el_vinculo_sin_duplicar_producto_componente_ni_vinculo()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        await e.PrepararAsync(id);
        var r = await e.PrepararAsync(id);
        await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Configurado, r.Estado);
        Assert.Equal(1, e.Api.ProductosCreados); Assert.Equal(1, e.Api.ComponentesCreados);
        Assert.Single(await e.VinculosAsync());
    }

    [Fact]
    public async Task Sin_proyecto_de_squash_no_aplica_ni_crea_nada()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Manual", null);
        var r = await e.PrepararAsync(id);
        Assert.False(r.Aplica);
        Assert.Empty(e.Api.Productos); Assert.Empty(await e.VinculosAsync());
    }

    [Fact]
    public async Task Fallo_al_crear_producto_conserva_la_iniciativa_y_el_reintento_lo_crea()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        e.Api.FallarProducto = true;

        var r1 = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Error, r1.Estado);
        var v = Assert.Single(await e.VinculosAsync());
        Assert.Equal(EstadoBugzilla.Error, v.Estado); Assert.Null(v.ProductoId); Assert.False(string.IsNullOrWhiteSpace(v.UltimoError));
        using (var db = e.NuevoContexto()) Assert.NotNull(await db.Proyectos.FindAsync(id));

        e.Api.FallarProducto = false;
        var r2 = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Configurado, r2.Estado);
        Assert.Equal(1, e.Api.ProductosCreados);
        Assert.Null((await e.VinculosAsync()).Single().UltimoError);
    }

    [Fact]
    public async Task Fallo_parcial_en_componente_guarda_el_producto_y_el_reintento_continua_sin_recrearlo()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        e.Api.FallarComponente = true;

        var r1 = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Error, r1.Estado);
        var parcial = Assert.Single(await e.VinculosAsync());
        Assert.NotNull(parcial.ProductoId); Assert.Null(parcial.ComponenteId);

        e.Api.FallarComponente = false;
        var r2 = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Configurado, r2.Estado);
        Assert.Equal(1, e.Api.ProductosCreados); Assert.Equal(1, e.Api.ComponentesCreados);
        var v = Assert.Single(await e.VinculosAsync());
        Assert.Equal(parcial.ProductoId, v.ProductoId); Assert.NotNull(v.ComponenteId);
    }

    [Fact]
    public async Task Producto_con_el_mismo_nombre_sin_marca_no_se_asocia_hasta_que_un_administrador_confirma()
    {
        await using var e = await Entorno.CrearAsync();
        e.Api.Productos.Add((7, "Proyecto Prueba 3", "Producto ajeno", new List<string> { "General" }, new List<string> { "1.0" }));
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);

        var r1 = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Error, r1.Estado);
        Assert.Null((await e.VinculosAsync()).Single().ProductoId);
        Assert.Equal(0, e.Api.ProductosCreados);

        var r2 = await e.PrepararAsync(id, confirmar: true);
        Assert.Equal(EstadoBugzilla.Configurado, r2.Estado);
        Assert.Equal(7, (await e.VinculosAsync()).Single().ProductoId);
        Assert.Equal(0, e.Api.ProductosCreados);
    }

    [Fact]
    public async Task Producto_con_la_marca_de_esta_iniciativa_se_adopta_tras_un_fallo_a_medias()
    {
        await using var e = await Entorno.CrearAsync();
        var marca = BugzillaProvision.Marca("http://localhost:8090/squash", 3);
        e.Api.Productos.Add((7, "Proyecto Prueba 3", "x " + marca, new List<string>(), new List<string> { "1.0" }));
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);

        var r = await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Configurado, r.Estado);
        Assert.Equal(0, e.Api.ProductosCreados); Assert.Equal(1, e.Api.ComponentesCreados);
    }

    [Fact]
    public async Task Un_producto_ya_vinculado_a_otra_iniciativa_no_se_reutiliza()
    {
        await using var e = await Entorno.CrearAsync();
        var a = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        await e.PrepararAsync(a);
        var b = await e.NuevoProyectoAsync("Proyecto Prueba 3", 4); // mismo nombre, otro proyecto de Squash
        var r = await e.PrepararAsync(b, confirmar: true);

        Assert.Equal(EstadoBugzilla.Error, r.Estado);
        Assert.Single(e.Api.Productos);
        Assert.Equal(1, (await e.VinculosAsync()).Count(v => v.ProductoId != null));
    }

    [Fact]
    public async Task Renombrar_la_iniciativa_mantiene_el_producto_vinculado()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        await e.PrepararAsync(id);
        var productoId = (await e.VinculosAsync()).Single().ProductoId;
        using (var db = e.NuevoContexto()) { (await db.Proyectos.FindAsync(id))!.Nombre = "Nombre nuevo"; await db.SaveChangesAsync(); }

        var r = await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Configurado, r.Estado);
        Assert.Equal(1, e.Api.ProductosCreados);
        Assert.Equal(productoId, (await e.VinculosAsync()).Single().ProductoId);
    }

    [Fact]
    public async Task Si_el_producto_vinculado_desaparece_no_se_crea_otro_automaticamente()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        await e.PrepararAsync(id);
        e.Api.Productos.Clear();

        var r = await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Error, r.Estado);
        Assert.Equal(1, e.Api.ProductosCreados);
        Assert.NotNull((await e.VinculosAsync()).Single().ProductoId);
    }

    [Fact]
    public async Task Version_ausente_sin_endpoint_de_version_queda_en_error_con_instruccion_manual()
    {
        await using var e = await Entorno.CrearAsync();
        e.Api.SinVersionPorApi = true;
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        var r = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Error, r.Estado);
        Assert.Contains("1.0", r.Mensaje);
        Assert.NotNull((await e.VinculosAsync()).Single().ComponenteId);
    }

    [Fact]
    public async Task Ambiente_de_squash_distinto_no_reutiliza_el_vinculo_ni_lo_reasigna()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        await e.PrepararAsync(id);
        var original = (await e.VinculosAsync()).Single();

        e.Squash = "https://squash.remoto.example/squash"; // mismo SquashProjectId = 3, otro servidor
        var r = await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Error, r.Estado);
        var v = Assert.Single(await e.VinculosAsync());
        Assert.Equal(original.ProductoId, v.ProductoId);
        Assert.Equal(original.SquashAmbiente, v.SquashAmbiente);
        Assert.Equal(1, e.Api.ProductosCreados);
    }

    [Fact]
    public async Task Ambiente_de_bugzilla_distinto_no_reutiliza_ids_de_otro_servidor()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        await e.PrepararAsync(id);
        var original = (await e.VinculosAsync()).Single();

        e.Bugzilla = "https://bugzilla.remoto.example";
        var r = await e.PrepararAsync(id);

        Assert.Equal(EstadoBugzilla.Error, r.Estado);
        var v = Assert.Single(await e.VinculosAsync());
        Assert.Equal(original.ProductoId, v.ProductoId);
        Assert.Equal("http://localhost:8080", v.BugzillaAmbiente);
    }

    [Fact]
    public async Task Mismo_id_numerico_en_otro_ambiente_es_otra_relacion_y_en_el_mismo_ambiente_se_rechaza()
    {
        await using var e = await Entorno.CrearAsync();
        var a = await e.NuevoProyectoAsync("Proyecto A", 5);
        await e.PrepararAsync(a);
        var b = await e.NuevoProyectoAsync("Proyecto B", 6);
        var otroSquash = "https://squash.remoto.example/squash";
        using (var db = e.NuevoContexto())
        {
            db.VinculosBugzilla.Add(new VinculoBugzilla { ProyectoId = b, SquashAmbiente = otroSquash, SquashProjectId = 5, BugzillaAmbiente = "https://bugzilla.remoto.example", ProductoId = 100 });
            await db.SaveChangesAsync();
        }
        Assert.Equal(2, (await e.VinculosAsync()).Count);

        var c = await e.NuevoProyectoAsync("Proyecto C", 7);
        using var db2 = e.NuevoContexto();
        db2.VinculosBugzilla.Add(new VinculoBugzilla { ProyectoId = c, SquashAmbiente = otroSquash, SquashProjectId = 5, BugzillaAmbiente = "x" });
        await Assert.ThrowsAsync<DbUpdateException>(() => db2.SaveChangesAsync());
    }

    [Fact]
    public async Task Solicitudes_concurrentes_no_duplican_producto_componente_ni_vinculo()
    {
        await using var e = await Entorno.CrearAsync();
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        e.Api.Demora = TimeSpan.FromMilliseconds(150);

        var resultados = await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() => e.PrepararAsync(id))));

        Assert.Equal(1, e.Api.ProductosCreados); Assert.Equal(1, e.Api.ComponentesCreados);
        Assert.Single(e.Api.Productos);
        Assert.Single(await e.VinculosAsync());
        Assert.Contains(resultados, r => r.Estado == EstadoBugzilla.Configurado);
        e.Api.Demora = TimeSpan.Zero;
        Assert.Equal(EstadoBugzilla.Configurado, (await e.PrepararAsync(id)).Estado);
        Assert.Equal(1, e.Api.ProductosCreados);
    }

    [Fact]
    public async Task Sin_cuenta_de_servicio_queda_pendiente_sin_llamar_a_bugzilla()
    {
        await using var e = await Entorno.CrearAsync();
        e.AdminKey = null;
        var id = await e.NuevoProyectoAsync("Proyecto Prueba 3", 3);
        var r = await e.PrepararAsync(id);
        Assert.Equal(EstadoBugzilla.Pendiente, r.Estado);
        Assert.Empty(e.Api.Productos);
        Assert.Equal(EstadoBugzilla.Pendiente, (await e.VinculosAsync()).Single().Estado);
    }
}

public class BugzillaProductosClientTests
{
    private sealed class Manejador(Func<HttpRequestMessage, HttpResponseMessage> f) : HttpMessageHandler
    {
        public List<(HttpMethod Metodo, string Ruta, string? Cuerpo)> Llamadas { get; } = new();
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            var cuerpo = request.Content is null ? null : await request.Content.ReadAsStringAsync(ct);
            Llamadas.Add((request.Method, request.RequestUri!.PathAndQuery, cuerpo));
            return f(request);
        }
    }

    private static HttpResponseMessage Json(string json, HttpStatusCode c = HttpStatusCode.OK) => new(c) { Content = new StringContent(json, Encoding.UTF8, "application/json") };
    private static BugzillaProductosClient Cliente(Manejador m) => new(new HttpClient(m) { BaseAddress = new Uri("http://localhost:8080/rest/") });

    [Fact]
    public async Task Crea_producto_y_componente_con_los_endpoints_oficiales()
    {
        var m = new Manejador(_ => Json("{\"id\": 42}"));
        var c = Cliente(m);
        Assert.Equal(42, (await c.CrearProductoAsync("P", "d", "1.0")).Datos);
        Assert.Equal(42, (await c.CrearComponenteAsync("P", "General", "d", "svc@x")).Datos);
        Assert.Equal("/rest/product", m.Llamadas[0].Ruta);
        Assert.Contains("\"version\":\"1.0\"", m.Llamadas[0].Cuerpo);
        Assert.Equal("/rest/component", m.Llamadas[1].Ruta);
        Assert.Contains("\"default_assignee\":\"svc@x\"", m.Llamadas[1].Cuerpo);
    }

    [Fact]
    public async Task Interpreta_producto_con_componentes_y_versiones()
    {
        var m = new Manejador(_ => Json("{\"products\":[{\"id\":9,\"name\":\"P\",\"description\":\"d\",\"components\":[{\"id\":1,\"name\":\"General\"}],\"versions\":[{\"id\":2,\"name\":\"1.0\"}]}]}"));
        var p = (await Cliente(m).ObtenerProductoAsync(9)).Datos!;
        Assert.Equal("General", Assert.Single(p.Componentes).Nombre);
        Assert.Equal("1.0", Assert.Single(p.Versiones).Nombre);
    }

    [Fact]
    public async Task Errores_de_permisos_y_de_bugzilla_no_exponen_credenciales_ni_html()
    {
        var m = new Manejador(r => r.Method == HttpMethod.Get ? Json("{}", HttpStatusCode.Unauthorized) : Json("{\"error\":true,\"code\":51,\"message\":\"<b>El producto</b>   ya existe\"}", HttpStatusCode.BadRequest));
        var c = Cliente(m);
        var a = await c.ListarProductosAsync();
        Assert.False(a.Ok); Assert.Contains("permisos", a.Error);
        var b = await c.CrearProductoAsync("P", "d", "1.0");
        Assert.False(b.Ok); Assert.Equal("Bugzilla: El producto ya existe", b.Error);
    }
}
