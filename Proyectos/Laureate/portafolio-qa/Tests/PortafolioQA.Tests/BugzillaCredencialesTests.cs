using System.Net;
using System.Text;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using PortafolioQA.Services.Conexion;
using PortafolioQA.Services.Integrations;

namespace PortafolioQA.Tests;

public class BugzillaCredencialesStoreTests : IDisposable
{
    private readonly string _ruta = Path.Combine(Path.GetTempPath(), "pqa-test-" + Guid.NewGuid().ToString("N"));
    private readonly IDataProtectionProvider _dp = new EphemeralDataProtectionProvider();

    private BugzillaCredencialesStore Crear(BugzillaOptions? servidor = null)
    {
        var config = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?> { ["Conexion:Ruta"] = _ruta }).Build();
        var conexion = new ConexionStore(config, _dp, NullLogger<ConexionStore>.Instance);
        return new BugzillaCredencialesStore(conexion, Options.Create(servidor ?? new BugzillaOptions { BaseUrl = "http://b" }), _dp,
            NullLogger<BugzillaCredencialesStore>.Instance);
    }

    public void Dispose() { if (Directory.Exists(_ruta)) Directory.Delete(_ruta, true); }

    [Fact]
    public void SinNadaGuardado_UsaLaConfiguracionDelServidor()
    {
        var s = Crear(new BugzillaOptions { BaseUrl = "http://b", AdminApiKey = "del-servidor", CuentaServicio = "srv@x" });
        Assert.Equal(OrigenBugzilla.Servidor, s.Origen);
        Assert.Equal("del-servidor", s.Actual.AdminApiKey);
        Assert.True(s.Actual.ProductosConfigurado);
    }

    [Fact]
    public void SinNadaConfigurado_QuedaPendiente()
    {
        var s = Crear();
        Assert.Equal(OrigenBugzilla.SinConfigurar, s.Origen);
        Assert.False(s.Actual.ProductosConfigurado);
    }

    [Fact]
    public void Guardar_TienePrioridad_SeActivaDeInmediato_YNoQuedaEnTextoPlano()
    {
        var s = Crear(new BugzillaOptions { BaseUrl = "http://b", AdminApiKey = "del-servidor" });
        s.Guardar(null, "  clave-secreta-123 ", " svc@x ", null, null, "admin@y");

        Assert.Equal(OrigenBugzilla.Aplicacion, s.Origen);
        Assert.Equal("clave-secreta-123", s.Actual.AdminApiKey);
        Assert.Equal("svc@x", s.Actual.CuentaServicio);
        Assert.Equal("admin@y", s.GuardadoPor);
        Assert.DoesNotContain("clave-secreta-123", File.ReadAllText(s.Archivo, Encoding.UTF8));
        Assert.DoesNotContain("clave-secreta-123", File.ReadAllText(Path.Combine(_ruta, "auditoria-configuracion.log")));
    }

    [Fact]
    public void Guardado_SobreviveAUnReinicio()
    {
        Crear().Guardar("http://nuevo/", "k1", "svc@x", "lec1", null, "admin");
        var otro = Crear();
        Assert.Equal("k1", otro.Actual.AdminApiKey);
        Assert.Equal("lec1", otro.Actual.ApiKey);
        Assert.Equal("http://nuevo", otro.Actual.BaseUrl);
        Assert.Equal(OrigenBugzilla.Aplicacion, otro.Origen);
    }

    [Fact]
    public void GuardarParcial_ConservaLoDemas_YLaDireccionGuardadaGanaAlServidor()
    {
        var s = Crear(new BugzillaOptions { BaseUrl = "http://servidor", ApiKey = "lectura-servidor" });
        s.Guardar("http://otro/rest/", "svc-key", "svc@x", null, null, "admin");   // solo servicio + dirección
        Assert.Equal("http://otro", s.Actual.BaseUrl);                       // sin /rest ni barra final
        Assert.Equal("lectura-servidor", s.Actual.ApiKey);                   // lectura: sigue la del servidor
        s.Guardar(null, null, null, "lec-nueva", null, "admin");                   // solo lectura: lo demás se conserva
        Assert.Equal("http://otro", s.Actual.BaseUrl);
        Assert.Equal("svc-key", s.Actual.AdminApiKey);
        Assert.Equal("lec-nueva", s.Actual.ApiKey);
        Assert.DoesNotContain("lec-nueva", File.ReadAllText(Path.Combine(_ruta, "auditoria-configuracion.log")));
    }

    [Fact]
    public void ArchivoDeLaVersionAnterior_SinDireccionNiLectura_SigueLeyendose()
    {
        // Formato guardado por la primera versión de la pantalla: solo clave de servicio, cuenta, quién y cuándo.
        var s = Crear(new BugzillaOptions { BaseUrl = "http://servidor", ApiKey = "lectura-servidor" });
        var proteger = _dp.CreateProtector("PortafolioQA.BugzillaServicio.v1");
        Directory.CreateDirectory(_ruta);
        File.WriteAllText(s.Archivo, proteger.Protect("{\"ApiKey\":\"vieja\",\"CuentaServicio\":\"svc@x\",\"GuardadoPor\":\"a\",\"GuardadoEn\":\"2026-10-19T10:00:00\"}"));
        s.Recargar();
        Assert.Equal("vieja", s.Actual.AdminApiKey);
        Assert.Equal("http://servidor", s.Actual.BaseUrl);
        Assert.Equal("lectura-servidor", s.Actual.ApiKey);
    }

    [Theory]
    [InlineData("http://b/", "http://b")]
    [InlineData(" https://b.pe/rest ", "https://b.pe")]
    [InlineData("https://b.pe/bugzilla/rest/", "https://b.pe/bugzilla")]
    [InlineData("http://10.0.0.1:8080/bugzilla/rest.cgi", "http://10.0.0.1:8080/bugzilla")]
    public void NormalizarUrl_DejaLaRaiz(string entrada, string esperado) =>
        Assert.Equal(esperado, BugzillaCredencialesStore.NormalizarUrl(entrada));

    [Fact]
    public void RutaRest_SeGuarda_YLasOpcionesArmanLaUrl()
    {
        var s = Crear(new BugzillaOptions { BaseUrl = "http://servidor" });
        Assert.Equal("http://servidor/rest/", s.Actual.UrlRest);                 // por defecto
        s.Guardar("http://10.0.0.1:8080/bugzilla", "k", "svc@x", null, "rest.cgi", "admin");
        Assert.Equal("http://10.0.0.1:8080/bugzilla/rest.cgi/", Crear().Actual.UrlRest);
        Assert.Equal("http://10.0.0.1:8080/bugzilla/rest.cgi/", s.Actual.UrlRest);
    }

    [Fact]
    public void Quitar_VuelveALaConfiguracionDelServidor()
    {
        var s = Crear(new BugzillaOptions { BaseUrl = "http://b", AdminApiKey = "del-servidor" });
        s.Guardar(null, "k1", "svc@x", null, null, "admin");
        s.Quitar("admin");
        Assert.Equal(OrigenBugzilla.Servidor, s.Origen);
        Assert.Equal("del-servidor", s.Actual.AdminApiKey);
    }

    [Fact]
    public void ArchivoIlegible_NoRompe_YAvisa()
    {
        var s = Crear();
        s.Guardar(null, "k1", "svc@x", null, null, "admin");
        File.WriteAllText(s.Archivo, "basura");
        s.Recargar();
        Assert.NotNull(s.Aviso);
        Assert.Equal(OrigenBugzilla.SinConfigurar, s.Origen);
    }
}

public class BugzillaProbadorTests
{
    private sealed class Falso(Func<string, (HttpStatusCode, string)> responde) : HttpMessageHandler
    {
        public List<string> Urls { get; } = new();
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage r, CancellationToken ct)
        {
            Urls.Add(r.RequestUri!.ToString());
            var (codigo, cuerpo) = responde(r.RequestUri.ToString());
            return Task.FromResult(new HttpResponseMessage(codigo) { Content = new StringContent(cuerpo, Encoding.UTF8, "application/json") });
        }
    }

    private static BugzillaProbador Con(Falso f) => new(new HttpClient(f));
    private const string Version = "{\"version\":\"5.2\"}";
    private const string Url = "http://bz";

    [Fact]
    public async Task TodoCorrecto_Aceptable_ConPermiso()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : u.Contains("valid_login") ? (HttpStatusCode.OK, "{\"result\":true}")
            : (HttpStatusCode.OK, "{\"users\":[{\"groups\":[{\"name\":\"editbugs\"},{\"name\":\"editcomponents\"}]}]}"));
        var r = await Con(f).ProbarAsync(Url + "/rest/", "K", "svc@x");   // acepta la dirección con /rest y la normaliza
        Assert.True(r.Aceptable);
        Assert.All(f.Urls, u => Assert.StartsWith("http://bz/rest/", u));
        Assert.True(r.TienePermiso);
        Assert.All(f.Urls.Where(u => !u.Contains("/version")), u => Assert.Contains("Bugzilla_api_key=K", u));
    }

    [Fact]
    public async Task ClaveInvalida_NoAceptable_YElMensajeNoLaRepite()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : (HttpStatusCode.OK, "{\"error\":true,\"code\":306,\"message\":\"The API key you specified is invalid.\"}"));
        var r = await Con(f).ProbarAsync(Url, "CLAVE-SECRETA", "svc@x");
        Assert.False(r.Aceptable);
        Assert.False(r.ClaveValida);
        Assert.DoesNotContain("CLAVE-SECRETA", r.Mensaje);
    }

    [Fact]
    public async Task ClaveDeOtraCuenta_NoAceptable()
    {
        var r = await Con(new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : (HttpStatusCode.OK, "{\"result\":false}"))).ProbarAsync(Url, "K", "svc@x");
        Assert.True(r.ClaveValida);
        Assert.False(r.CuentaCoincide);
        Assert.False(r.Aceptable);
    }

    [Fact]
    public async Task SinEditcomponents_NoAceptable()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : u.Contains("valid_login") ? (HttpStatusCode.OK, "{\"result\":true}")
            : (HttpStatusCode.OK, "{\"users\":[{\"groups\":[{\"name\":\"editbugs\"}]}]}"));
        var r = await Con(f).ProbarAsync(Url, "K", "svc@x");
        Assert.False(r.TienePermiso);
        Assert.False(r.Aceptable);
    }

    [Fact]
    public async Task GruposNoInformados_EsAceptablePeroNoVerificado()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : u.Contains("valid_login") ? (HttpStatusCode.OK, "{\"result\":true}") : (HttpStatusCode.OK, "{\"users\":[{\"name\":\"svc@x\"}]}"));
        var r = await Con(f).ProbarAsync(Url, "K", "svc@x");
        Assert.Null(r.TienePermiso);
        Assert.True(r.Aceptable);
    }

    [Theory]
    [InlineData(null, "K", "svc@x")]
    [InlineData("no-es-url", "K", "svc@x")]
    [InlineData(Url, "", "svc@x")]
    [InlineData(Url, "K", " ")]
    public async Task EntradasIncompletas_NoLlamanABugzilla(string? url, string clave, string cuenta)
    {
        var f = new Falso(_ => (HttpStatusCode.OK, "{}"));
        var r = await Con(f).ProbarAsync(url, clave, cuenta);
        Assert.False(r.Aceptable);
        Assert.Empty(f.Urls);
    }

    [Fact]
    public async Task DireccionQueNoEsBugzilla_NoAceptable_YNoSeEnviaLaClave()
    {
        var f = new Falso(_ => (HttpStatusCode.OK, "<html>otra cosa</html>"));
        var r = await Con(f).ProbarAsync(Url, "CLAVE-SECRETA", "svc@x");
        Assert.False(r.Aceptable);
        Assert.Contains("no responde como Bugzilla", r.Mensaje);
        Assert.Equal(2, f.Urls.Count);                          // solo /rest/version y /rest.cgi/version: la clave no sale hacia un servidor desconocido
        Assert.All(f.Urls, u => Assert.DoesNotContain("CLAVE-SECRETA", u));
    }

    [Fact]
    public async Task SinReescritura_DetectaRestCgi_YUsaEsaRuta()
    {
        // Como el Apache del servidor de QA: /bugzilla/rest/... da 404 y solo /bugzilla/rest.cgi/... responde.
        var f = new Falso(u => u.Contains("/rest.cgi/version") ? (HttpStatusCode.OK, Version)
            : u.Contains("/rest/") ? (HttpStatusCode.NotFound, "<html>404</html>")
            : u.Contains("valid_login") ? (HttpStatusCode.OK, "{\"result\":true}")
            : (HttpStatusCode.OK, "{\"users\":[{\"groups\":[{\"name\":\"editcomponents\"}]}]}"));
        var r = await Con(f).ProbarAsync("http://10.0.0.1:8080/bugzilla/", "K", "svc@x");
        Assert.True(r.Aceptable);
        Assert.Equal("rest.cgi", r.RutaRest);
        Assert.Contains(f.Urls, u => u.StartsWith("http://10.0.0.1:8080/bugzilla/rest.cgi/valid_login?"));
    }

    [Fact]
    public async Task ConReescritura_LaRutaEsRest()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : (HttpStatusCode.OK, "{\"products\":[]}"));
        Assert.Equal("rest", (await Con(f).ProbarLecturaAsync(Url, "LEC")).RutaRest);
    }

    [Fact]
    public async Task Lectura_ClaveReconocida_Aceptable()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : (HttpStatusCode.OK, "{\"products\":[]}"));
        var r = await Con(f).ProbarLecturaAsync(Url, "LEC");
        Assert.True(r.Aceptable);
        Assert.Contains(f.Urls, u => u.Contains("product?") && u.Contains("Bugzilla_api_key=LEC"));
    }

    [Fact]
    public async Task Lectura_ClaveInvalida_NoAceptable()
    {
        var f = new Falso(u => u.Contains("/version") ? (HttpStatusCode.OK, Version) : (HttpStatusCode.OK, "{\"error\":true,\"code\":306,\"message\":\"invalid\"}"));
        var r = await Con(f).ProbarLecturaAsync(Url, "LEC");
        Assert.False(r.Aceptable);
        Assert.DoesNotContain("LEC", r.Mensaje);
    }

    [Fact]
    public async Task BugzillaCaido_MensajeClaro()
    {
        var p = new BugzillaProbador(new HttpClient(new Lanza()));
        var r = await p.ProbarAsync(Url, "K", "svc@x");
        Assert.False(r.Conecta);
        Assert.False(r.Aceptable);
    }

    private sealed class Lanza : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage r, CancellationToken ct) => throw new HttpRequestException("caído");
    }

    [Fact]
    public async Task Handler_AnadeLaClaveComoParametro_ConYSinQuery()
    {
        var capturadas = new List<string>();
        var interno = new Falso(u => { capturadas.Add(u); return (HttpStatusCode.OK, "{}"); });
        var clave = "a";
        var handler = new BugzillaApiKeyHandler(() => clave) { InnerHandler = interno };
        using var http = new HttpClient(handler) { BaseAddress = new Uri("http://bz/rest/") };
        await http.GetAsync("product");
        clave = "b&c";                         // la clave se lee en cada llamada: un cambio surte efecto sin reiniciar
        await http.GetAsync("bug?product=X");
        Assert.EndsWith("product?Bugzilla_api_key=a", capturadas[0]);
        Assert.EndsWith("bug?product=X&Bugzilla_api_key=b%26c", capturadas[1]);
    }
}
